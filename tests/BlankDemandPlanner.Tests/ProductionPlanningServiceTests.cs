using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Services.Planning;
using FluentAssertions;

namespace BlankDemandPlanner.Tests;

public sealed class ProductionPlanningServiceTests
{
    private readonly ProductionPlanningService service = new();

    [Fact]
    public void BuildPlan_keeps_route_order_and_prevents_resource_overlap()
    {
        var start = new DateTime(2026, 8, 3, 8, 0, 0);
        var demands = new[]
        {
            new PlanningDemand(1, "Заказ 1", "1816772", "Вал", 2, start.AddDays(5), 50),
            new PlanningDemand(2, "Заказ 2", "1816772", "Вал", 1, start.AddDays(6), 10)
        };
        var routes = new[]
        {
            new PlanningRoute(10, "1816772", 10, "010", "Токарная", "Токарные", "Токарь", 3, 10, 0, 30, 5, false),
            new PlanningRoute(11, "1816772", 20, "020", "Шлифовальная", "Шлифовальные", "Шлифовщик", 3, 5, 0, 20, 5, false)
        };
        var equipment = new[]
        {
            new PlanningEquipmentResource(100, "T-1", "Токарный 1", "Токарные", 1m, 1m, true),
            new PlanningEquipmentResource(200, "S-1", "Шлифовальный 1", "Шлифовальные", 1m, 1m, true)
        };
        var employees = new[]
        {
            new PlanningEmployeeResource(1000, "1", "Иванов", "Токарь", 4, 8, 17, 40, true),
            new PlanningEmployeeResource(2000, "2", "Петров", "Шлифовщик", 4, 8, 17, 40, true)
        };

        var result = service.BuildPlan(demands, routes, equipment, employees, start);

        result.Issues.Should().BeEmpty();
        result.Operations.Should().HaveCount(4);
        var firstOrder = result.Operations.Where(x => x.DemandItemId == 1).OrderBy(x => x.PlannedStart).ToArray();
        firstOrder[1].PlannedStart.Should().BeOnOrAfter(firstOrder[0].PlannedEnd);
        var turning = result.Operations.Where(x => x.EquipmentId == 100).OrderBy(x => x.PlannedStart).ToArray();
        turning[1].PlannedStart.Should().BeOnOrAfter(turning[0].PlannedEnd);
    }

    [Fact]
    public void BuildPlan_reports_missing_qualification_instead_of_assigning_employee()
    {
        var start = new DateTime(2026, 8, 3, 8, 0, 0);
        var result = service.BuildPlan(
            [new PlanningDemand(1, "Заказ", "1816772", "Вал", 1, start.AddDays(2), 50)],
            [new PlanningRoute(10, "1816772", 10, "010", "Токарная", "Токарные", "Токарь", 5, 0, 10, 0, 0, false)],
            [new PlanningEquipmentResource(100, "T-1", "Токарный 1", "Токарные", 1m, 1m, true)],
            [new PlanningEmployeeResource(1000, "1", "Иванов", "Токарь", 4, 8, 17, 40, true)],
            start);

        result.Operations.Should().BeEmpty();
        result.Issues.Should().ContainSingle(x => x.Message.Contains("разряд 5+"));
    }

    [Fact]
    public void BuildPlan_applies_equipment_efficiency_factor()
    {
        var start = new DateTime(2026, 8, 3, 8, 0, 0);
        var route = new PlanningRoute(10, "1816772", 10, "010", "Токарная", "Токарные", "Токарь", 3, 0, 0, 60, 0, false);
        var employee = new PlanningEmployeeResource(1000, "1", "Иванов", "Токарь", 4, 8, 17, 40, true);

        var result = service.BuildPlan(
            [new PlanningDemand(1, "Заказ", "1816772", "Вал", 1, start.AddDays(2), 50)],
            [route],
            [new PlanningEquipmentResource(100, "T-1", "Токарный 1", "Токарные", 1m, 0.8m, true)],
            [employee],
            start);

        (result.Operations.Single().PlannedEnd - start).TotalMinutes.Should().Be(75);
    }
}
