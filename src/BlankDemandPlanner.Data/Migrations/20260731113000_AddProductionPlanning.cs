using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260731113000_AddProductionPlanning")]
public partial class AddProductionPlanning : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
CREATE TABLE IF NOT EXISTS "ProductionEquipment" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ProductionEquipment" PRIMARY KEY AUTOINCREMENT,
    "Code" TEXT NOT NULL,
    "Name" TEXT NOT NULL,
    "Model" TEXT NULL,
    "ResourceGroup" TEXT NOT NULL,
    "CapacityPerHour" TEXT NOT NULL,
    "EfficiencyFactor" TEXT NOT NULL,
    "Status" TEXT NOT NULL,
    "MaintenanceSchedule" TEXT NULL,
    "IsActive" INTEGER NOT NULL,
    "UpdatedAt" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_ProductionEquipment_Code" ON "ProductionEquipment" ("Code");
CREATE INDEX IF NOT EXISTS "IX_ProductionEquipment_ResourceGroup" ON "ProductionEquipment" ("ResourceGroup");

CREATE TABLE IF NOT EXISTS "ProductionEmployees" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ProductionEmployees" PRIMARY KEY AUTOINCREMENT,
    "PersonnelNumber" TEXT NOT NULL,
    "FullName" TEXT NOT NULL,
    "Specialty" TEXT NOT NULL,
    "QualificationLevel" INTEGER NOT NULL,
    "ShiftStartHour" INTEGER NOT NULL,
    "ShiftEndHour" INTEGER NOT NULL,
    "MaxHoursPerWeek" TEXT NOT NULL,
    "IsAvailable" INTEGER NOT NULL,
    "AbsenceReason" TEXT NULL,
    "UpdatedAt" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_ProductionEmployees_PersonnelNumber" ON "ProductionEmployees" ("PersonnelNumber");
CREATE INDEX IF NOT EXISTS "IX_ProductionEmployees_Specialty" ON "ProductionEmployees" ("Specialty");

CREATE TABLE IF NOT EXISTS "ProductionRouteOperations" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ProductionRouteOperations" PRIMARY KEY AUTOINCREMENT,
    "Ips" TEXT NOT NULL,
    "Sequence" INTEGER NOT NULL,
    "OperationCode" TEXT NOT NULL,
    "Description" TEXT NOT NULL,
    "EquipmentGroup" TEXT NOT NULL,
    "RequiredSpecialty" TEXT NOT NULL,
    "MinimumQualification" INTEGER NOT NULL,
    "SetupMinutes" TEXT NOT NULL,
    "PieceMinutes" TEXT NOT NULL,
    "MachineMinutes" TEXT NOT NULL,
    "AuxiliaryMinutes" TEXT NOT NULL,
    "CanRunInParallel" INTEGER NOT NULL,
    "UpdatedAt" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_ProductionRouteOperations_Ips_Sequence" ON "ProductionRouteOperations" ("Ips", "Sequence");
CREATE INDEX IF NOT EXISTS "IX_ProductionRouteOperations_Ips" ON "ProductionRouteOperations" ("Ips");

CREATE TABLE IF NOT EXISTS "ProductionScheduleEntries" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ProductionScheduleEntries" PRIMARY KEY AUTOINCREMENT,
    "DemandItemId" INTEGER NOT NULL,
    "RouteOperationId" INTEGER NOT NULL,
    "EquipmentId" INTEGER NOT NULL,
    "EmployeeId" INTEGER NOT NULL,
    "OrderNumber" TEXT NOT NULL,
    "Ips" TEXT NOT NULL,
    "PartName" TEXT NOT NULL,
    "OperationCode" TEXT NOT NULL,
    "OperationName" TEXT NOT NULL,
    "EquipmentName" TEXT NOT NULL,
    "EmployeeName" TEXT NOT NULL,
    "Quantity" TEXT NOT NULL,
    "Priority" INTEGER NOT NULL,
    "PlannedStart" TEXT NOT NULL,
    "PlannedEnd" TEXT NOT NULL,
    "DueDate" TEXT NOT NULL,
    "Status" TEXT NOT NULL,
    "CreatedAt" TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS "IX_ProductionScheduleEntries_PlannedStart" ON "ProductionScheduleEntries" ("PlannedStart");
CREATE INDEX IF NOT EXISTS "IX_ProductionScheduleEntries_EquipmentId" ON "ProductionScheduleEntries" ("EquipmentId");
CREATE INDEX IF NOT EXISTS "IX_ProductionScheduleEntries_EmployeeId" ON "ProductionScheduleEntries" ("EmployeeId");
CREATE INDEX IF NOT EXISTS "IX_ProductionScheduleEntries_DemandItemId" ON "ProductionScheduleEntries" ("DemandItemId");
""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
DROP TABLE IF EXISTS "ProductionScheduleEntries";
DROP TABLE IF EXISTS "ProductionRouteOperations";
DROP TABLE IF EXISTS "ProductionEmployees";
DROP TABLE IF EXISTS "ProductionEquipment";
""");
    }
}
