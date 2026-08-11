namespace BlankDemandPlanner.Core.Models;

public sealed record PlanningDemand(
    long DemandItemId,
    string OrderNumber,
    string Ips,
    string PartName,
    decimal Quantity,
    DateTime DueDate,
    int Priority);

public sealed record PlanningRoute(
    long Id,
    string Ips,
    int Sequence,
    string OperationCode,
    string Description,
    string EquipmentGroup,
    string RequiredSpecialty,
    int MinimumQualification,
    decimal SetupMinutes,
    decimal PieceMinutes,
    decimal MachineMinutes,
    decimal AuxiliaryMinutes,
    bool CanRunInParallel);

public sealed record PlanningEquipmentResource(
    long Id,
    string Code,
    string Name,
    string ResourceGroup,
    decimal CapacityPerHour,
    decimal EfficiencyFactor,
    bool IsAvailable);

public sealed record PlanningEmployeeResource(
    long Id,
    string PersonnelNumber,
    string FullName,
    string Specialty,
    int QualificationLevel,
    int ShiftStartHour,
    int ShiftEndHour,
    decimal MaxHoursPerWeek,
    bool IsAvailable);

public sealed record PlannedOperation(
    long DemandItemId,
    long RouteOperationId,
    long EquipmentId,
    long EmployeeId,
    string OrderNumber,
    string Ips,
    string PartName,
    string OperationCode,
    string OperationName,
    string EquipmentName,
    string EmployeeName,
    decimal Quantity,
    int Priority,
    DateTime PlannedStart,
    DateTime PlannedEnd,
    DateTime DueDate,
    string Status);

public sealed record PlanningIssue(string Ips, string OperationCode, string Message);

public sealed record ProductionPlanningResult(
    IReadOnlyList<PlannedOperation> Operations,
    IReadOnlyList<PlanningIssue> Issues);
