using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Services.Planning;
using FluentAssertions;

namespace BlankDemandPlanner.Tests;

public sealed class PzmcPersonnelAvailabilityServiceTests
{
    [Fact]
    public async Task Availability_filters_by_configured_mechanical_entity()
    {
        var snapshot = new PzmcProductionPersonnelSnapshot(
            [
                new PzmcEmployee { Id = 1, EntityId = 44, FullName = "Иванов И.И.", Position = "Токарь", Schedule = "1 смена" },
                new PzmcEmployee { Id = 2, EntityId = 10, FullName = "Петров П.П.", Position = "Слесарь", Schedule = "1 смена" }
            ],
            [],
            [new PzmcSkudLogUser { EmployeeId = 1, DateTime = DateTime.Today.AddHours(8), EventName = "Вход" }],
            [new PzmcEntitie { Id = 44, Name = "Механообработка" }],
            DateTime.Now,
            "Тест");
        var service = new PzmcPersonnelAvailabilityService(new FakeClient(snapshot, 44));

        var result = await service.GetAvailabilityAsync(true, CancellationToken.None);

        result.Rows.Should().ContainSingle();
        result.Rows[0].FullName.Should().Be("Иванов И.И.");
        result.Rows[0].AvailableToday.Should().BeTrue();
        result.Rows[0].Entity.Should().Be("Механообработка");
    }

    [Fact]
    public async Task Availability_uses_skud_as_capacity_layer()
    {
        var snapshot = new PzmcProductionPersonnelSnapshot(
            [new PzmcEmployee { Id = 1, EntityId = 44, FullName = "Иванов И.И.", Position = "Токарь", PlannedHours = 8 }],
            [new PzmcEmployeeStatistic { EmployeeId = 1, Date = DateTime.Today, ProductionHours = 2 }],
            [
                new PzmcSkudLogUser { EmployeeId = 1, DateTime = DateTime.Today.AddHours(8), EventName = "Вход" },
                new PzmcSkudLogUser { EmployeeId = 1, DateTime = DateTime.Today.AddHours(12), EventName = "Проход" }
            ],
            [],
            DateTime.Now,
            "Тест");
        var service = new PzmcPersonnelAvailabilityService(new FakeClient(snapshot, 44));

        var result = await service.GetAvailabilityAsync(true, CancellationToken.None);

        result.Rows.Should().ContainSingle();
        result.Rows[0].SkudHours.Should().Be(4);
        result.Rows[0].ProductionHours.Should().Be(2);
        result.Rows[0].AvailableHoursToday.Should().Be(6);
    }

    [Fact]
    public async Task Availability_maps_actual_production_employee_payload()
    {
        var snapshot = new PzmcProductionPersonnelSnapshot(
            [
                new PzmcEmployee
                {
                    Id = 271,
                    UserId = 384,
                    EntityId = 431,
                    LastName = "Ivanov",
                    FirstName = "Ivan",
                    MiddleName = "Ivanovich",
                    EmployeePost = new PzmcEmployeePost { Name = "Operator" },
                    EmployeeDepartment = new PzmcEntityRef { Id = 431, Name = "Machining workshop" },
                    ScheduleId = 1
                }
            ],
            [new PzmcEmployeeStatistic { EmployeeId = 271, DatePlan = DateTime.Today, AvailableHours = 8.2, UsedHours = 1.2 }],
            [new PzmcSkudLogUser { UserId = 384, Time = DateTime.Today.AddHours(9), DirectionName = "in" }],
            [new PzmcEntitie { Id = 431, EntityName = "Machining workshop" }],
            DateTime.Now,
            "Test");
        var service = new PzmcPersonnelAvailabilityService(new FakeClient(snapshot, 431));

        var result = await service.GetAvailabilityAsync(true, CancellationToken.None);

        result.Rows.Should().ContainSingle();
        result.Rows[0].FullName.Should().Be("Ivanov Ivan Ivanovich");
        result.Rows[0].Position.Should().Be("Operator");
        result.Rows[0].Entity.Should().Be("Machining workshop");
        result.Rows[0].PlannedHours.Should().Be(8.2m);
        result.Rows[0].ProductionHours.Should().Be(1.2m);
        result.Rows[0].AvailableToday.Should().BeTrue();
    }

    private sealed class FakeClient(PzmcProductionPersonnelSnapshot snapshot, int? mechanicalEntityId) : IPzmcProductionApiClient
    {
        public bool IsConfigured => true;
        public IReadOnlyList<string> CmoGroups => [];
        public int? MechanicalEntityId => mechanicalEntityId;
        public int PersonnelRefreshMinutes => 10;
        public Task<PzmcProductionSnapshot> LoadSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<PzmcProductionPersonnelSnapshot> LoadPersonnelSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken) => Task.FromResult(snapshot);
        public Task<IReadOnlyList<PzmcProductNeed>> GetProductNeedsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcProductNeed>>([]);
        public Task<IReadOnlyList<PzmcProduct>> GetProductsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcProduct>>([]);
        public Task<IReadOnlyList<PzmcSpecIpsObject>> GetIpsObjectsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcSpecIpsObject>>([]);
        public Task<IReadOnlyList<PzmcSklad>> GetWarehousesAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcSklad>>([]);
        public Task<IReadOnlyList<PzmcSkladPlanBase>> GetPlannedReceiptsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcSkladPlanBase>>([]);
        public Task<IReadOnlyList<PzmcAnalog>> GetAnalogsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcAnalog>>([]);
        public Task<IReadOnlyList<PzmcProductGroup>> GetProductGroupsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcProductGroup>>([]);
        public Task<IReadOnlyList<PzmcEmployee>> GetEmployeesAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcEmployee>>([]);
        public Task<IReadOnlyList<PzmcEmployeeStatistic>> GetEmployeeStatisticsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcEmployeeStatistic>>([]);
        public Task<IReadOnlyList<PzmcSkudLogUser>> GetSkudLogUsersAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcSkudLogUser>>([]);
        public Task<IReadOnlyList<PzmcEntitie>> GetEntitiesAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcEntitie>>([]);
    }
}
