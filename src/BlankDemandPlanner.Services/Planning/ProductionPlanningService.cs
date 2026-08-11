using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Services.Planning;

public sealed class ProductionPlanningService : IProductionPlanningService
{
    public ProductionPlanningResult BuildPlan(
        IReadOnlyCollection<PlanningDemand> demands,
        IReadOnlyCollection<PlanningRoute> routes,
        IReadOnlyCollection<PlanningEquipmentResource> equipment,
        IReadOnlyCollection<PlanningEmployeeResource> employees,
        DateTime planningStart)
    {
        var operations = new List<PlannedOperation>();
        var issues = new List<PlanningIssue>();
        var equipmentFreeAt = new Dictionary<long, DateTime>();
        var employeeFreeAt = new Dictionary<long, DateTime>();
        var employeeWeeklyMinutes = new Dictionary<(long EmployeeId, DateTime WeekStart), decimal>();

        var routesByIps = routes
            .GroupBy(x => x.Ips.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderBy(r => r.Sequence).ThenBy(r => r.Id).ToArray(), StringComparer.OrdinalIgnoreCase);

        foreach (var demand in demands
                     .Where(x => x.Quantity > 0)
                     .OrderByDescending(x => x.Priority)
                     .ThenBy(x => x.DueDate)
                     .ThenBy(x => x.OrderNumber, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(x => x.Ips, StringComparer.OrdinalIgnoreCase))
        {
            if (!routesByIps.TryGetValue(demand.Ips.Trim(), out var demandRoutes) || demandRoutes.Length == 0)
            {
                issues.Add(new PlanningIssue(demand.Ips, string.Empty, "Для детали не задан технологический маршрут."));
                continue;
            }

            var predecessorEnd = planningStart;
            foreach (var route in demandRoutes)
            {
                var compatibleEquipment = equipment
                    .Where(x => x.IsAvailable && Same(x.ResourceGroup, route.EquipmentGroup))
                    .ToArray();
                var compatibleEmployees = employees
                    .Where(x => x.IsAvailable &&
                                Same(x.Specialty, route.RequiredSpecialty) &&
                                x.QualificationLevel >= route.MinimumQualification)
                    .ToArray();

                if (compatibleEquipment.Length == 0)
                {
                    issues.Add(new PlanningIssue(demand.Ips, route.OperationCode, $"Нет доступного оборудования группы «{route.EquipmentGroup}»."));
                    break;
                }

                if (compatibleEmployees.Length == 0)
                {
                    issues.Add(new PlanningIssue(demand.Ips, route.OperationCode, $"Нет доступного сотрудника: {route.RequiredSpecialty}, разряд {route.MinimumQualification}+."));
                    break;
                }

                Candidate? best = null;
                foreach (var machine in compatibleEquipment)
                {
                    foreach (var employee in compatibleEmployees)
                    {
                        var earliest = Max(
                            planningStart,
                            predecessorEnd,
                            equipmentFreeAt.GetValueOrDefault(machine.Id, planningStart),
                            employeeFreeAt.GetValueOrDefault(employee.Id, planningStart));
                        var durationMinutes = CalculateDurationMinutes(route, demand.Quantity, machine.CapacityPerHour, machine.EfficiencyFactor);
                        var start = AlignToWeeklyCapacity(
                            earliest,
                            durationMinutes,
                            employee,
                            employeeWeeklyMinutes);
                        var end = AddWorkingMinutes(start, durationMinutes, employee.ShiftStartHour, employee.ShiftEndHour);
                        var candidate = new Candidate(machine, employee, start, end, durationMinutes);
                        if (best is null || candidate.End < best.End ||
                            (candidate.End == best.End && candidate.Start < best.Start))
                        {
                            best = candidate;
                        }
                    }
                }

                if (best is null)
                {
                    issues.Add(new PlanningIssue(demand.Ips, route.OperationCode, "Не удалось найти свободное окно времени."));
                    break;
                }

                equipmentFreeAt[best.Equipment.Id] = best.End;
                employeeFreeAt[best.Employee.Id] = best.End;
                var weekKey = (best.Employee.Id, GetWeekStart(best.Start));
                employeeWeeklyMinutes[weekKey] = employeeWeeklyMinutes.GetValueOrDefault(weekKey) + best.DurationMinutes;
                predecessorEnd = route.CanRunInParallel ? predecessorEnd : best.End;

                operations.Add(new PlannedOperation(
                    demand.DemandItemId,
                    route.Id,
                    best.Equipment.Id,
                    best.Employee.Id,
                    demand.OrderNumber,
                    demand.Ips,
                    demand.PartName,
                    route.OperationCode,
                    route.Description,
                    best.Equipment.Name,
                    best.Employee.FullName,
                    demand.Quantity,
                    demand.Priority,
                    best.Start,
                    best.End,
                    demand.DueDate,
                    best.End > demand.DueDate ? "Риск срыва" : "Запланировано"));
            }
        }

        return new ProductionPlanningResult(
            operations.OrderBy(x => x.PlannedStart).ThenBy(x => x.EquipmentName).ToArray(),
            issues);
    }

    private static decimal CalculateDurationMinutes(
        PlanningRoute route,
        decimal quantity,
        decimal capacityPerHour,
        decimal efficiencyFactor)
    {
        var efficiency = efficiencyFactor <= 0 ? 1m : efficiencyFactor;
        var perPiece = route.PieceMinutes + route.MachineMinutes + route.AuxiliaryMinutes;
        if (perPiece <= 0 && capacityPerHour > 0)
        {
            perPiece = 60m / capacityPerHour;
        }

        return Math.Max(1m, route.SetupMinutes + perPiece * quantity) / efficiency;
    }

    private static DateTime AlignToWeeklyCapacity(
        DateTime earliest,
        decimal durationMinutes,
        PlanningEmployeeResource employee,
        IReadOnlyDictionary<(long EmployeeId, DateTime WeekStart), decimal> weeklyMinutes)
    {
        var candidate = AlignToShift(earliest, employee.ShiftStartHour, employee.ShiftEndHour);
        var weeklyLimit = Math.Max(1m, employee.MaxHoursPerWeek) * 60m;
        for (var attempts = 0; attempts < 52; attempts++)
        {
            var key = (employee.Id, GetWeekStart(candidate));
            if (weeklyMinutes.GetValueOrDefault(key) + Math.Min(durationMinutes, weeklyLimit) <= weeklyLimit)
            {
                return candidate;
            }

            candidate = GetWeekStart(candidate).AddDays(7).AddHours(employee.ShiftStartHour);
        }

        return candidate;
    }

    private static DateTime GetWeekStart(DateTime value)
    {
        var offset = ((int)value.DayOfWeek + 6) % 7;
        return value.Date.AddDays(-offset);
    }

    private static DateTime AddWorkingMinutes(DateTime start, decimal durationMinutes, int shiftStartHour, int shiftEndHour)
    {
        var current = AlignToShift(start, shiftStartHour, shiftEndHour);
        var remaining = (double)durationMinutes;
        while (remaining > 0)
        {
            var shiftEnd = current.Date.AddHours(shiftEndHour);
            var available = Math.Max(0, (shiftEnd - current).TotalMinutes);
            if (remaining <= available)
            {
                return current.AddMinutes(remaining);
            }

            remaining -= available;
            current = AlignToShift(current.Date.AddDays(1).AddHours(shiftStartHour), shiftStartHour, shiftEndHour);
        }

        return current;
    }

    private static DateTime AlignToShift(DateTime value, int shiftStartHour, int shiftEndHour)
    {
        var startHour = Math.Clamp(shiftStartHour, 0, 23);
        var endHour = Math.Clamp(shiftEndHour, startHour + 1, 24);
        var current = value;
        while (true)
        {
            if (current.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday)
            {
                current = current.Date.AddDays(current.DayOfWeek == DayOfWeek.Saturday ? 2 : 1).AddHours(startHour);
                continue;
            }

            var shiftStart = current.Date.AddHours(startHour);
            var shiftEnd = current.Date.AddHours(endHour);
            if (current < shiftStart)
            {
                return shiftStart;
            }

            if (current >= shiftEnd)
            {
                current = current.Date.AddDays(1).AddHours(startHour);
                continue;
            }

            return current;
        }
    }

    private static bool Same(string left, string right) =>
        string.Equals(left.Trim(), right.Trim(), StringComparison.OrdinalIgnoreCase);

    private static DateTime Max(params DateTime[] values) => values.Max();

    private sealed record Candidate(
        PlanningEquipmentResource Equipment,
        PlanningEmployeeResource Employee,
        DateTime Start,
        DateTime End,
        decimal DurationMinutes);
}
