using System.Text.Json.Serialization;
using System.Text.Json;

namespace BlankDemandPlanner.Core.Models;

public sealed class PzmcProductNeed
{
    [JsonPropertyName("products_need_id")] public int Id { get; set; }
    [JsonPropertyName("product_id")] public int ProductId { get; set; }
    [JsonPropertyName("technology_id")] public int? TechnologyId { get; set; }
    [JsonPropertyName("operation_id")] public int? OperationId { get; set; }
    [JsonPropertyName("designation")] public string Designation { get; set; } = string.Empty;
    [JsonPropertyName("op_name")] public string OperationName { get; set; } = string.Empty;
    [JsonPropertyName("unit")] public string Unit { get; set; } = string.Empty;
    [JsonPropertyName("qty")] public double Quantity { get; set; }
    [JsonPropertyName("qty_exists")] public double QuantityExists { get; set; }
    [JsonPropertyName("need_date")] public DateTime NeedDate { get; set; }
    [JsonPropertyName("doc_osnovanie_id")] public int? OrderId { get; set; }
    [JsonPropertyName("doc_osnovanie_link")] public string OrderLink { get; set; } = string.Empty;
    [JsonPropertyName("doc_osnovanie_number")] public int? OrderNumber { get; set; }
    [JsonPropertyName("doc_osnovanie_type")] public string OrderType { get; set; } = string.Empty;
    [JsonPropertyName("object_id")] public int ObjectId { get; set; }
    [JsonPropertyName("object_ips_id")] public int ObjectIpsId { get; set; }
    [JsonPropertyName("object_name")] public string ObjectName { get; set; } = string.Empty;
    [JsonPropertyName("object_designation")] public string ObjectDesignation { get; set; } = string.Empty;
    [JsonPropertyName("manuf_type")] public string ManufacturingType { get; set; } = string.Empty;
    [JsonPropertyName("manuf_method")] public string ManufacturingMethod { get; set; } = string.Empty;
    [JsonPropertyName("manager_1c")] public string Manager { get; set; } = string.Empty;
    [JsonPropertyName("supply_1c")] public string Supplier { get; set; } = string.Empty;
    [JsonPropertyName("product_group_id")] public int? ProductGroupId { get; set; }
    [JsonPropertyName("class_name")] public string ClassName { get; set; } = string.Empty;
    [JsonPropertyName("pdf_exists")] public bool PdfExists { get; set; }
}

public sealed class PzmcProduct
{
    [JsonPropertyName("product_id")] public int Id { get; set; }
    [JsonPropertyName("serial_id")] public string SerialId { get; set; } = string.Empty;
    [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("priority")] public int Priority { get; set; }
    [JsonPropertyName("project_name")] public string ProjectName { get; set; } = string.Empty;
    [JsonPropertyName("technology_id")] public int TechnologyId { get; set; }
    [JsonPropertyName("need_date")] public DateTime? NeedDate { get; set; }
    [JsonPropertyName("doc_osnovanie_id")] public int? OrderId { get; set; }
}

public sealed class PzmcSpecIpsObject
{
    [JsonPropertyName("object_id")] public int ObjectId { get; set; }
    [JsonPropertyName("object_ips_id")] public int ObjectIpsId { get; set; }
    [JsonPropertyName("object_name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("object_designation")] public string Designation { get; set; } = string.Empty;
    [JsonPropertyName("manuf_type")] public string ManufacturingType { get; set; } = string.Empty;
    [JsonPropertyName("manuf_method")] public string ManufacturingMethod { get; set; } = string.Empty;
    [JsonPropertyName("manager_1c")] public string Manager { get; set; } = string.Empty;
    [JsonPropertyName("supply_1c")] public string Supplier { get; set; } = string.Empty;
    [JsonPropertyName("product_group_id")] public int? ProductGroupId { get; set; }
    [JsonPropertyName("class_name")] public string ClassName { get; set; } = string.Empty;
    [JsonPropertyName("pdf_exists")] public bool PdfExists { get; set; }
}

public sealed class PzmcSklad
{
    [JsonPropertyName("sklad_id")] public int Id { get; set; }
    [JsonPropertyName("designation")] public string Designation { get; set; } = string.Empty;
    [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("is_special")] public bool IsSpecial { get; set; }
    [JsonPropertyName("is_used_in_production")] public bool IsUsedInProduction { get; set; }
    [JsonPropertyName("Balance")] public List<PzmcSkladBalance> Balance { get; set; } = [];
}

public sealed class PzmcSkladBalance
{
    [JsonPropertyName("sklad_balance_id")] public int Id { get; set; }
    [JsonPropertyName("sklad_id")] public int SkladId { get; set; }
    [JsonPropertyName("object_id")] public int ObjectId { get; set; }
    [JsonPropertyName("qty")] public double Quantity { get; set; }
    [JsonPropertyName("qty_reserve")] public double ReservedQuantity { get; set; }
    [JsonPropertyName("qty_remaining")] public double RemainingQuantity { get; set; }
    [JsonPropertyName("unit")] public string Unit { get; set; } = string.Empty;
}

public sealed class PzmcSkladPlanBase
{
    [JsonPropertyName("sklad_plan_base_id")] public int Id { get; set; }
    [JsonPropertyName("sklad_plan_type")] public string PlanType { get; set; } = string.Empty;
    [JsonPropertyName("object_id")] public int ObjectId { get; set; }
    [JsonPropertyName("qty")] public double Quantity { get; set; }
    [JsonPropertyName("unit")] public string Unit { get; set; } = string.Empty;
    [JsonPropertyName("date_plan")] public DateTime PlannedDate { get; set; }
    [JsonPropertyName("date_actual")] public DateTime? ActualDate { get; set; }
    [JsonPropertyName("order_ref_key")] public string OrderReference { get; set; } = string.Empty;
    [JsonPropertyName("project")] public string Project { get; set; } = string.Empty;
    [JsonPropertyName("supply_name")] public string Supplier { get; set; } = string.Empty;
}

public sealed class PzmcAnalog
{
    [JsonPropertyName("analog_id")] public int Id { get; set; }
    [JsonPropertyName("object_parent_id")] public int ParentObjectId { get; set; }
    [JsonPropertyName("object_id")] public int ObjectId { get; set; }
    [JsonPropertyName("date_start")] public DateTime? DateStart { get; set; }
    [JsonPropertyName("date_end")] public DateTime? DateEnd { get; set; }
    [JsonPropertyName("is_priority")] public bool IsPriority { get; set; }
    [JsonPropertyName("tech_id")] public int[] TechnologyIds { get; set; } = [];
    [JsonPropertyName("products_id")] public int[] ProductIds { get; set; } = [];
}

public sealed class PzmcProductGroup
{
    [JsonPropertyName("product_group_id")] public int Id { get; set; }
    [JsonPropertyName("designation")] public string Designation { get; set; } = string.Empty;
    [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("sort_id")] public int SortId { get; set; }
}

public sealed class PzmcEmployee
{
    [JsonPropertyName("employee_id")] public int Id { get; set; }
    [JsonPropertyName("user_id")] public int UserId { get; set; }
    [JsonPropertyName("entity_id")] public int? EntityId { get; set; }
    [JsonPropertyName("post_id")] public int? PostId { get; set; }
    [JsonPropertyName("fio")] public string FullName { get; set; } = string.Empty;
    [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("user_last_name")] public string LastName { get; set; } = string.Empty;
    [JsonPropertyName("user_first_name")] public string FirstName { get; set; } = string.Empty;
    [JsonPropertyName("user_middle_name")] public string MiddleName { get; set; } = string.Empty;
    [JsonPropertyName("position")] public string Position { get; set; } = string.Empty;
    [JsonPropertyName("post_name")] public string PostName { get; set; } = string.Empty;
    [JsonPropertyName("employee_post")] public PzmcEmployeePost? EmployeePost { get; set; }
    [JsonPropertyName("employee_department")] public PzmcEntityRef? EmployeeDepartment { get; set; }
    [JsonPropertyName("employer")] public PzmcEntityRef? Employer { get; set; }
    [JsonPropertyName("schedule")] public string Schedule { get; set; } = string.Empty;
    [JsonPropertyName("schedule_id")] public int? ScheduleId { get; set; }
    [JsonPropertyName("plan_hours")] public double? PlannedHours { get; set; }
    [JsonPropertyName("is_active")] public bool? IsActive { get; set; }
    [JsonExtensionData] public Dictionary<string, JsonElement> Extra { get; set; } = [];
}

public sealed class PzmcEmployeePost
{
    [JsonPropertyName("post_id")] public int Id { get; set; }
    [JsonPropertyName("post_name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("post_short")] public string ShortName { get; set; } = string.Empty;
    [JsonPropertyName("entity_id")] public int? EntityId { get; set; }
}

public sealed class PzmcEntityRef
{
    [JsonPropertyName("entity_id")] public int Id { get; set; }
    [JsonPropertyName("entity_pid")] public int? ParentId { get; set; }
    [JsonPropertyName("entity_name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("is_production_place")] public bool IsProductionPlace { get; set; }
    [JsonPropertyName("is_main")] public bool IsMain { get; set; }
}

public sealed class PzmcEmployeeStatistic
{
    [JsonPropertyName("employee_statistic_id")] public int Id { get; set; }
    [JsonPropertyName("entity_employees_statistic_id")] public int EntityEmployeesStatisticId { get; set; }
    [JsonPropertyName("employee_id")] public int EmployeeId { get; set; }
    [JsonPropertyName("entity_id")] public int? EntityId { get; set; }
    [JsonPropertyName("date")] public DateTime? Date { get; set; }
    [JsonPropertyName("date_plan")] public DateTime? DatePlan { get; set; }
    [JsonPropertyName("plan_hours")] public double? PlannedHours { get; set; }
    [JsonPropertyName("skud_hours")] public double? SkudHours { get; set; }
    [JsonPropertyName("production_hours")] public double? ProductionHours { get; set; }
    [JsonPropertyName("available")] public double? AvailableHours { get; set; }
    [JsonPropertyName("used")] public double? UsedHours { get; set; }
    [JsonPropertyName("plan_user")] public double? UserPlannedHours { get; set; }
    [JsonPropertyName("plan_tech")] public double? TechnologyPlannedHours { get; set; }
    [JsonPropertyName("note")] public string Note { get; set; } = string.Empty;
    [JsonExtensionData] public Dictionary<string, JsonElement> Extra { get; set; } = [];
}

public sealed class PzmcSkudLogUser
{
    [JsonPropertyName("skud_log_user_id")] public int Id { get; set; }
    [JsonPropertyName("employee_id")] public int EmployeeId { get; set; }
    [JsonPropertyName("user_id")] public int UserId { get; set; }
    [JsonPropertyName("entity_id")] public int? EntityId { get; set; }
    [JsonPropertyName("time")] public DateTime? Time { get; set; }
    [JsonPropertyName("date_time")] public DateTime? DateTime { get; set; }
    [JsonPropertyName("event_name")] public string EventName { get; set; } = string.Empty;
    [JsonPropertyName("direction")] public string Direction { get; set; } = string.Empty;
    [JsonPropertyName("dir_name")] public string DirectionName { get; set; } = string.Empty;
    [JsonPropertyName("dir_desc")] public string DirectionDescription { get; set; } = string.Empty;
    [JsonExtensionData] public Dictionary<string, JsonElement> Extra { get; set; } = [];
}

public sealed class PzmcEntitie
{
    [JsonPropertyName("entity_id")] public int Id { get; set; }
    [JsonPropertyName("entity_pid")] public int? EntityParentId { get; set; }
    [JsonPropertyName("parent_id")] public int? ParentId { get; set; }
    [JsonPropertyName("entity_name")] public string EntityName { get; set; } = string.Empty;
    [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("designation")] public string Designation { get; set; } = string.Empty;
    [JsonExtensionData] public Dictionary<string, JsonElement> Extra { get; set; } = [];
}

public sealed record PzmcProductionSnapshot(
    IReadOnlyList<PzmcProductNeed> Needs,
    IReadOnlyList<PzmcProduct> Products,
    IReadOnlyList<PzmcSpecIpsObject> IpsObjects,
    IReadOnlyList<PzmcSklad> Warehouses,
    IReadOnlyList<PzmcSkladPlanBase> PlannedReceipts,
    IReadOnlyList<PzmcAnalog> Analogs,
    IReadOnlyList<PzmcProductGroup> ProductGroups,
    DateTime UpdatedAt,
    string Source);

public sealed record PzmcProductionPersonnelSnapshot(
    IReadOnlyList<PzmcEmployee> Employees,
    IReadOnlyList<PzmcEmployeeStatistic> EmployeeStatistics,
    IReadOnlyList<PzmcSkudLogUser> SkudLogs,
    IReadOnlyList<PzmcEntitie> Entities,
    DateTime UpdatedAt,
    string Source);

public sealed record PzmcNeedRow(
    int NeedId,
    string Project,
    string SerialNumber,
    int Priority,
    int Ips,
    string Designation,
    string Name,
    string ProductGroup,
    string ManufacturingType,
    string ManufacturingMethod,
    string Manager,
    string Supplier,
    string Unit,
    decimal RequiredQuantity,
    decimal StockQuantity,
    decimal AnalogStockQuantity,
    decimal PlannedReceiptQuantity,
    decimal ShortageQuantity,
    DateTime NeedDate,
    bool HasPdf,
    bool HasOrder,
    string Order,
    string Operation);

public sealed record PzmcNeedFilters(
    string Project = "",
    string ProductGroup = "",
    string ManufacturingType = "",
    string ManufacturingMethod = "",
    string Manager = "",
    string Supplier = "",
    DateTime? DueFrom = null,
    DateTime? DueTo = null,
    bool? HasPdf = null,
    bool? HasOrder = null,
    bool CmoOnly = true);

public sealed record PzmcNeedResult(
    IReadOnlyList<PzmcNeedRow> Rows,
    DateTime UpdatedAt,
    string Source,
    string Status);

public sealed record PzmcPersonnelAvailabilityRow(
    int EmployeeId,
    int? EntityId,
    string Entity,
    string FullName,
    string Position,
    string Schedule,
    decimal PlannedHours,
    decimal SkudHours,
    decimal ProductionHours,
    decimal AvailableHoursToday,
    bool AvailableToday,
    string Note);

public sealed record PzmcPersonnelAvailabilityResult(
    IReadOnlyList<PzmcPersonnelAvailabilityRow> Rows,
    DateTime UpdatedAt,
    string Source,
    string Status);

public static class PzmcDynamicJson
{
    public static string GetString(IReadOnlyDictionary<string, JsonElement> values, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (!values.TryGetValue(key, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.String)
            {
                return value.GetString() ?? string.Empty;
            }

            if (value.ValueKind is JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False)
            {
                return value.ToString();
            }
        }

        return string.Empty;
    }

    public static int? GetInt(IReadOnlyDictionary<string, JsonElement> values, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (!values.TryGetValue(key, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            {
                return number;
            }

            if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out number))
            {
                return number;
            }
        }

        return null;
    }

    public static decimal? GetDecimal(IReadOnlyDictionary<string, JsonElement> values, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (!values.TryGetValue(key, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.Number && value.TryGetDecimal(out var number))
            {
                return number;
            }

            if (value.ValueKind == JsonValueKind.String &&
                decimal.TryParse(value.GetString()?.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number))
            {
                return number;
            }
        }

        return null;
    }

    public static DateTime? GetDateTime(IReadOnlyDictionary<string, JsonElement> values, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (!values.TryGetValue(key, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.String && DateTime.TryParse(value.GetString(), out var parsed))
            {
                return parsed;
            }
        }

        return null;
    }
}
