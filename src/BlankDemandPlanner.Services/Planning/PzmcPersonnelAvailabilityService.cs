using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Services.Planning;

public sealed class PzmcPersonnelAvailabilityService(IPzmcProductionApiClient apiClient) : IPzmcPersonnelAvailabilityService
{
    public async Task<PzmcPersonnelAvailabilityResult> GetAvailabilityAsync(bool forceRefresh, CancellationToken cancellationToken)
    {
        var snapshot = await apiClient.LoadPersonnelSnapshotAsync(forceRefresh, cancellationToken);
        var rows = Calculate(snapshot)
            .OrderByDescending(x => x.AvailableToday)
            .ThenBy(x => x.Position)
            .ThenBy(x => x.FullName)
            .ToArray();

        return new PzmcPersonnelAvailabilityResult(
            rows,
            snapshot.UpdatedAt,
            snapshot.Source,
            $"Персонал: {rows.Length}; на месте: {rows.Count(x => x.AvailableToday)}; источник: {snapshot.Source}; обновлено {snapshot.UpdatedAt:dd.MM.yyyy HH:mm}");
    }

    private IReadOnlyList<PzmcPersonnelAvailabilityRow> Calculate(PzmcProductionPersonnelSnapshot snapshot)
    {
        var today = DateTime.Today;
        var entities = snapshot.Entities
            .Select(x => new
            {
                Id = ResolveEntityId(x),
                Name = First(x.EntityName, x.Name, x.Designation, PzmcDynamicJson.GetString(x.Extra, "entity_name", "caption", "title"))
            })
            .Where(x => x.Id > 0)
            .GroupBy(x => x.Id)
            .ToDictionary(x => x.Key, x => x.First().Name);

        var statisticsByEmployee = snapshot.EmployeeStatistics
            .Where(x => ResolveStatisticEmployeeId(x) > 0 && IsToday(ResolveStatisticDate(x), today))
            .GroupBy(ResolveStatisticEmployeeId)
            .ToDictionary(x => x.Key, x => x.ToArray());

        var employeeIdByUserId = snapshot.Employees
            .Where(x => ResolveEmployeeId(x) > 0 && ResolveEmployeeUserId(x) > 0)
            .GroupBy(ResolveEmployeeUserId)
            .ToDictionary(x => x.Key, x => ResolveEmployeeId(x.First()));
        var skudByEmployee = snapshot.SkudLogs
            .Select(x => new { EmployeeId = ResolveSkudEmployeeId(x, employeeIdByUserId), Log = x })
            .Where(x => x.EmployeeId > 0 && IsToday(ResolveSkudDateTime(x.Log), today))
            .GroupBy(x => x.EmployeeId)
            .ToDictionary(x => x.Key, x => x.Select(v => v.Log).OrderBy(ResolveSkudDateTime).ToArray());

        var rows = new List<PzmcPersonnelAvailabilityRow>();
        foreach (var employee in snapshot.Employees)
        {
            var employeeId = ResolveEmployeeId(employee);
            if (employeeId <= 0)
            {
                continue;
            }

            var entityId = ResolveEmployeeEntityId(employee);
            if (apiClient.MechanicalEntityId is not null && entityId != apiClient.MechanicalEntityId.Value)
            {
                continue;
            }

            if (ResolveEmployeeActive(employee) == false)
            {
                continue;
            }

            statisticsByEmployee.TryGetValue(employeeId, out var statistics);
            skudByEmployee.TryGetValue(employeeId, out var logs);

            var plannedHours = FirstDecimal(
                employee.PlannedHours is null ? null : (decimal)employee.PlannedHours.Value,
                SumStatistics(statistics, ResolvePlannedHours));
            var skudHours = FirstDecimal(
                SumStatistics(statistics, ResolveSkudHours),
                CalculateSkudHours(logs));
            var productionHours = SumStatistics(statistics, ResolveProductionHours) ?? 0m;
            var available = IsAvailable(logs, skudHours);
            var availableHours = available ? Math.Max(0m, plannedHours - productionHours) : 0m;
            var note = BuildNote(logs, statistics, apiClient.MechanicalEntityId, entityId);

            rows.Add(new PzmcPersonnelAvailabilityRow(
                employeeId,
                entityId,
                ResolveEmployeeEntityName(employee, entityId, entities),
                ResolveEmployeeFullName(employee),
                First(employee.Position, employee.PostName, employee.EmployeePost?.Name, PzmcDynamicJson.GetString(employee.Extra, "post", "dolzhnost", "position_name", "specialty")),
                First(employee.Schedule, ResolveSchedule(employee), PzmcDynamicJson.GetString(employee.Extra, "grafik", "work_schedule", "shift", "schedule_name")),
                plannedHours,
                skudHours,
                productionHours,
                availableHours,
                available,
                note));
        }

        return rows;
    }

    private static int ResolveEmployeeId(PzmcEmployee employee) =>
        employee.Id > 0
            ? employee.Id
            : PzmcDynamicJson.GetInt(employee.Extra, "id", "employeeId", "user_id", "userId") ?? 0;

    private static int ResolveEmployeeUserId(PzmcEmployee employee) =>
        employee.UserId > 0
            ? employee.UserId
            : PzmcDynamicJson.GetInt(employee.Extra, "userId", "user_id") ?? 0;

    private static int? ResolveEmployeeEntityId(PzmcEmployee employee) =>
        employee.EntityId ??
        (employee.EmployeeDepartment?.Id > 0 ? employee.EmployeeDepartment.Id : (int?)null) ??
        PzmcDynamicJson.GetInt(employee.Extra, "entityId", "department_id", "departmentId", "ceh_id", "workshop_id");

    private static string ResolveEmployeeEntityName(PzmcEmployee employee, int? entityId, IReadOnlyDictionary<int, string> entities) =>
        First(
            employee.EmployeeDepartment?.Name,
            entityId is not null && entities.TryGetValue(entityId.Value, out var entityName) ? entityName : string.Empty,
            employee.Employer?.Name);

    private static string ResolveEmployeeFullName(PzmcEmployee employee)
    {
        var combined = string.Join(" ", new[] { employee.LastName, employee.FirstName, employee.MiddleName }
            .Where(x => !string.IsNullOrWhiteSpace(x)));
        return First(
            employee.FullName,
            employee.Name,
            combined,
            PzmcDynamicJson.GetString(employee.Extra, "full_name", "fullname", "fio", "name", "employee_name"));
    }

    private static string ResolveSchedule(PzmcEmployee employee) =>
        employee.ScheduleId is null ? string.Empty : $"schedule_id {employee.ScheduleId.Value}";

    private static bool? ResolveEmployeeActive(PzmcEmployee employee)
    {
        if (employee.IsActive is not null)
        {
            return employee.IsActive;
        }

        var text = PzmcDynamicJson.GetString(employee.Extra, "active", "status", "is_working");
        if (string.IsNullOrWhiteSpace(text))
        {
            var dateTo = PzmcDynamicJson.GetDateTime(employee.Extra, "employee_date_to", "date_to", "dismissed_at", "deleted_at");
            return dateTo is null ? null : dateTo.Value.Date >= DateTime.Today;
        }

        return !text.Contains("увол", StringComparison.OrdinalIgnoreCase) &&
               !text.Equals("false", StringComparison.OrdinalIgnoreCase) &&
               !text.Equals("0", StringComparison.OrdinalIgnoreCase);
    }

    private static int ResolveStatisticEmployeeId(PzmcEmployeeStatistic statistic) =>
        statistic.EmployeeId > 0
            ? statistic.EmployeeId
            : PzmcDynamicJson.GetInt(statistic.Extra, "employeeId", "user_id", "userId") ?? 0;

    private static DateTime? ResolveStatisticDate(PzmcEmployeeStatistic statistic) =>
        statistic.Date ?? statistic.DatePlan ?? PzmcDynamicJson.GetDateTime(statistic.Extra, "work_date", "date_begin", "date_start", "created_at", "date_plan");

    private static decimal? ResolvePlannedHours(PzmcEmployeeStatistic statistic) =>
        FirstDecimalNullable(
            statistic.PlannedHours is null ? null : (decimal)statistic.PlannedHours.Value,
            statistic.AvailableHours is null ? null : (decimal)statistic.AvailableHours.Value,
            PzmcDynamicJson.GetDecimal(statistic.Extra, "planned_hours", "plan", "plan_time", "hours_plan", "available"));

    private static decimal? ResolveSkudHours(PzmcEmployeeStatistic statistic) =>
        statistic.SkudHours is null
            ? PzmcDynamicJson.GetDecimal(statistic.Extra, "fact_skud", "skud_time", "presence_hours", "hours_skud")
            : (decimal)statistic.SkudHours.Value;

    private static decimal? ResolveProductionHours(PzmcEmployeeStatistic statistic) =>
        FirstDecimalNullable(
            statistic.ProductionHours is null ? null : (decimal)statistic.ProductionHours.Value,
            statistic.UsedHours is null ? null : (decimal)statistic.UsedHours.Value,
            statistic.UserPlannedHours is null && statistic.TechnologyPlannedHours is null
                ? null
                : (decimal)((statistic.UserPlannedHours ?? 0d) + (statistic.TechnologyPlannedHours ?? 0d)),
            PzmcDynamicJson.GetDecimal(statistic.Extra, "production_time", "prod_hours", "work_hours", "hours_fact", "used"));

    private static int ResolveSkudEmployeeId(PzmcSkudLogUser log, IReadOnlyDictionary<int, int> employeeIdByUserId)
    {
        if (log.EmployeeId > 0)
        {
            return log.EmployeeId;
        }

        var userId = log.UserId > 0
            ? log.UserId
            : PzmcDynamicJson.GetInt(log.Extra, "userId", "user_id") ?? 0;
        return userId > 0 && employeeIdByUserId.TryGetValue(userId, out var employeeId) ? employeeId : 0;
    }

    private static DateTime? ResolveSkudDateTime(PzmcSkudLogUser log) =>
        log.DateTime ?? log.Time ?? PzmcDynamicJson.GetDateTime(log.Extra, "datetime", "date", "event_time", "created_at", "time");

    private static int ResolveEntityId(PzmcEntitie entity) =>
        entity.Id > 0
            ? entity.Id
            : PzmcDynamicJson.GetInt(entity.Extra, "id", "entityId") ?? 0;

    private static decimal? FirstDecimalNullable(params decimal?[] values) => values.FirstOrDefault(x => x is not null);

    private static decimal? SumStatistics(IEnumerable<PzmcEmployeeStatistic>? statistics, Func<PzmcEmployeeStatistic, decimal?> selector)
    {
        if (statistics is null)
        {
            return null;
        }

        var values = statistics.Select(selector).Where(x => x is not null).Select(x => x!.Value).ToArray();
        return values.Length == 0 ? null : values.Sum();
    }

    private static decimal? CalculateSkudHours(IReadOnlyList<PzmcSkudLogUser>? logs)
    {
        if (logs is null || logs.Count < 2)
        {
            return null;
        }

        var first = ResolveSkudDateTime(logs.First());
        var last = ResolveSkudDateTime(logs.Last());
        return first is null || last is null || last <= first
            ? null
            : (decimal)(last.Value - first.Value).TotalHours;
    }

    private static bool IsAvailable(IReadOnlyList<PzmcSkudLogUser>? logs, decimal skudHours)
    {
        if (logs is null || logs.Count == 0)
        {
            return false;
        }

        var last = logs.Last();
        var marker = First(last.Direction, last.DirectionName, last.DirectionDescription, last.EventName, PzmcDynamicJson.GetString(last.Extra, "event", "operation", "status"));
        if (marker.Contains("выход", StringComparison.OrdinalIgnoreCase) ||
            marker.Contains("out", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return skudHours >= 0;
    }

    private static string BuildNote(
        IReadOnlyList<PzmcSkudLogUser>? logs,
        IReadOnlyList<PzmcEmployeeStatistic>? statistics,
        int? configuredEntityId,
        int? entityId)
    {
        var parts = new List<string>();
        if (logs is null || logs.Count == 0)
        {
            parts.Add("нет прохода СКУД сегодня");
        }
        else
        {
            var last = ResolveSkudDateTime(logs.Last());
            if (last is not null)
            {
                parts.Add($"последний СКУД {last:HH:mm}");
            }
        }

        if (statistics is null || statistics.Count == 0)
        {
            parts.Add("нет статистики производства за сегодня");
        }

        if (configuredEntityId is null)
        {
            parts.Add("entity_id цеха не задан");
        }
        else if (entityId != configuredEntityId.Value)
        {
            parts.Add("не выбранный цех");
        }

        return string.Join("; ", parts);
    }

    private static bool IsToday(DateTime? value, DateTime today) => value is not null && value.Value.Date == today.Date;
    private static decimal FirstDecimal(params decimal?[] values) => values.FirstOrDefault(x => x is not null) ?? 0m;
    private static string First(params string?[] values) =>
        values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;
}
