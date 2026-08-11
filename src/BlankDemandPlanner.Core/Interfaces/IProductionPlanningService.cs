using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Core.Interfaces;

public interface IProductionPlanningService
{
    ProductionPlanningResult BuildPlan(
        IReadOnlyCollection<PlanningDemand> demands,
        IReadOnlyCollection<PlanningRoute> routes,
        IReadOnlyCollection<PlanningEquipmentResource> equipment,
        IReadOnlyCollection<PlanningEmployeeResource> employees,
        DateTime planningStart);
}
