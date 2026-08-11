using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Services.Planning;
using FluentAssertions;

namespace BlankDemandPlanner.Tests;

public sealed class PzmcNeedServiceTests
{
    [Fact]
    public async Task Deficit_allocates_stock_once_in_need_date_order()
    {
        var firstDate = new DateTime(2026, 8, 10);
        var snapshot = CreateSnapshot(
            needs:
            [
                new PzmcProductNeed { Id = 1, ProductId = 1, ObjectId = 100, ObjectIpsId = 1816772, Quantity = 4, Unit = "шт", NeedDate = firstDate },
                new PzmcProductNeed { Id = 2, ProductId = 1, ObjectId = 100, ObjectIpsId = 1816772, Quantity = 4, Unit = "шт", NeedDate = firstDate.AddDays(1) }
            ],
            stock: [new PzmcSkladBalance { ObjectId = 100, Quantity = 5, Unit = "шт" }]);
        var service = new PzmcNeedService(new FakeClient(snapshot, ["Заготовки"]));

        var result = await service.GetDeficitAsync(new PzmcNeedFilters(CmoOnly: true), true, CancellationToken.None);

        result.Rows.Should().ContainSingle();
        result.Rows[0].NeedId.Should().Be(2);
        result.Rows[0].StockQuantity.Should().Be(1);
        result.Rows[0].ShortageQuantity.Should().Be(3);
    }

    [Fact]
    public async Task Deficit_uses_applicable_analog_and_receipt_before_due_date()
    {
        var dueDate = new DateTime(2026, 8, 10);
        var snapshot = CreateSnapshot(
            needs:
            [
                new PzmcProductNeed { Id = 1, ProductId = 1, ObjectId = 100, ObjectIpsId = 1816772, Quantity = 6, Unit = "шт", NeedDate = dueDate }
            ],
            stock: [new PzmcSkladBalance { ObjectId = 200, Quantity = 2, Unit = "шт" }],
            plans: [new PzmcSkladPlanBase { ObjectId = 100, Quantity = 3, Unit = "шт", PlannedDate = dueDate.AddDays(-1) }],
            analogs: [new PzmcAnalog { Id = 1, ParentObjectId = 100, ObjectId = 200, IsPriority = true }]);
        var service = new PzmcNeedService(new FakeClient(snapshot, ["Заготовки"]));

        var result = await service.GetDeficitAsync(new PzmcNeedFilters(CmoOnly: true), true, CancellationToken.None);

        result.Rows.Should().ContainSingle();
        result.Rows[0].AnalogStockQuantity.Should().Be(2);
        result.Rows[0].PlannedReceiptQuantity.Should().Be(3);
        result.Rows[0].ShortageQuantity.Should().Be(1);
    }

    private static PzmcProductionSnapshot CreateSnapshot(
        IReadOnlyList<PzmcProductNeed> needs,
        IReadOnlyList<PzmcSkladBalance> stock,
        IReadOnlyList<PzmcSkladPlanBase>? plans = null,
        IReadOnlyList<PzmcAnalog>? analogs = null)
    {
        return new PzmcProductionSnapshot(
            needs,
            [new PzmcProduct { Id = 1, ProjectName = "ПС001", SerialId = "Станок 1", Priority = 50 }],
            [new PzmcSpecIpsObject { ObjectId = 100, ObjectIpsId = 1816772, Name = "Вал", ProductGroupId = 1 }],
            [new PzmcSklad { Id = 1, IsUsedInProduction = true, Balance = stock.ToList() }],
            plans ?? [],
            analogs ?? [],
            [new PzmcProductGroup { Id = 1, Name = "Заготовки ЦМО" }],
            DateTime.Now,
            "Тест");
    }

    private sealed class FakeClient(PzmcProductionSnapshot snapshot, IReadOnlyList<string> cmoGroups) : IPzmcProductionApiClient
    {
        public bool IsConfigured => true;
        public IReadOnlyList<string> CmoGroups => cmoGroups;
        public int? MechanicalEntityId => null;
        public int PersonnelRefreshMinutes => 10;
        public Task<PzmcProductionSnapshot> LoadSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken) => Task.FromResult(snapshot);
        public Task<PzmcProductionPersonnelSnapshot> LoadPersonnelSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken) => Task.FromResult(new PzmcProductionPersonnelSnapshot([], [], [], [], DateTime.Now, "РўРµСЃС‚"));
        public Task<IReadOnlyList<PzmcProductNeed>> GetProductNeedsAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.Needs);
        public Task<IReadOnlyList<PzmcProduct>> GetProductsAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.Products);
        public Task<IReadOnlyList<PzmcSpecIpsObject>> GetIpsObjectsAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.IpsObjects);
        public Task<IReadOnlyList<PzmcSklad>> GetWarehousesAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.Warehouses);
        public Task<IReadOnlyList<PzmcSkladPlanBase>> GetPlannedReceiptsAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.PlannedReceipts);
        public Task<IReadOnlyList<PzmcAnalog>> GetAnalogsAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.Analogs);
        public Task<IReadOnlyList<PzmcProductGroup>> GetProductGroupsAsync(CancellationToken cancellationToken) => Task.FromResult(snapshot.ProductGroups);
        public Task<IReadOnlyList<PzmcEmployee>> GetEmployeesAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcEmployee>>([]);
        public Task<IReadOnlyList<PzmcEmployeeStatistic>> GetEmployeeStatisticsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcEmployeeStatistic>>([]);
        public Task<IReadOnlyList<PzmcSkudLogUser>> GetSkudLogUsersAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcSkudLogUser>>([]);
        public Task<IReadOnlyList<PzmcEntitie>> GetEntitiesAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<PzmcEntitie>>([]);
    }
}
