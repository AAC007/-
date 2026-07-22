using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260715100000_InitialCreate")]
public partial class InitialCreate : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
CREATE TABLE IF NOT EXISTS "Parts" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_Parts" PRIMARY KEY AUTOINCREMENT,
    "Ips" TEXT NOT NULL,
    "Designation" TEXT NULL,
    "Name" TEXT NOT NULL,
    "HasMsk" INTEGER NOT NULL,
    "Source" TEXT NULL,
    "CreatedAt" TEXT NOT NULL,
    "UpdatedAt" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_Parts_Ips" ON "Parts" ("Ips");

CREATE TABLE IF NOT EXISTS "CanonicalBlanks" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_CanonicalBlanks" PRIMARY KEY AUTOINCREMENT,
    "CanonicalName" TEXT NOT NULL,
    "CanonicalKey" TEXT NOT NULL,
    "BlankType" INTEGER NOT NULL,
    "Material" TEXT NULL,
    "MaterialGost" TEXT NULL,
    "ProfileGost" TEXT NULL,
    "DiameterMm" TEXT NULL,
    "WidthMm" TEXT NULL,
    "HeightMm" TEXT NULL,
    "ThicknessMm" TEXT NULL,
    "WallThicknessMm" TEXT NULL,
    "LengthMm" TEXT NULL,
    "BaseUnit" INTEGER NOT NULL,
    "I012Status" INTEGER NOT NULL,
    "I012Section" TEXT NULL,
    "CreatedAt" TEXT NOT NULL,
    "UpdatedAt" TEXT NOT NULL,
    "IsActive" INTEGER NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_CanonicalBlanks_CanonicalKey" ON "CanonicalBlanks" ("CanonicalKey");

CREATE TABLE IF NOT EXISTS "BlankAliases" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_BlankAliases" PRIMARY KEY AUTOINCREMENT,
    "CanonicalBlankId" INTEGER NOT NULL,
    "OneCCode" TEXT NOT NULL,
    "SourceName" TEXT NOT NULL,
    "NormalizedSourceName" TEXT NOT NULL,
    "Source" TEXT NOT NULL,
    "ImportedAt" TEXT NOT NULL,
    "UpdatedAt" TEXT NOT NULL,
    "IsActive" INTEGER NOT NULL,
    CONSTRAINT "FK_BlankAliases_CanonicalBlanks_CanonicalBlankId" FOREIGN KEY ("CanonicalBlankId") REFERENCES "CanonicalBlanks" ("Id") ON DELETE CASCADE
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_BlankAliases_OneCCode" ON "BlankAliases" ("OneCCode");
CREATE INDEX IF NOT EXISTS "IX_BlankAliases_CanonicalBlankId" ON "BlankAliases" ("CanonicalBlankId");

CREATE TABLE IF NOT EXISTS "PartBlankMaps" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_PartBlankMaps" PRIMARY KEY AUTOINCREMENT,
    "PartId" INTEGER NOT NULL,
    "CanonicalBlankId" INTEGER NOT NULL,
    "ConsumptionQuantity" TEXT NOT NULL,
    "ConsumptionUnit" INTEGER NOT NULL,
    "LossPercent" TEXT NOT NULL,
    "Source" TEXT NOT NULL,
    "SourceFile" TEXT NULL,
    "UpdatedAt" TEXT NOT NULL,
    "IsActive" INTEGER NOT NULL,
    "IsPrimary" INTEGER NOT NULL,
    CONSTRAINT "FK_PartBlankMaps_CanonicalBlanks_CanonicalBlankId" FOREIGN KEY ("CanonicalBlankId") REFERENCES "CanonicalBlanks" ("Id") ON DELETE CASCADE,
    CONSTRAINT "FK_PartBlankMaps_Parts_PartId" FOREIGN KEY ("PartId") REFERENCES "Parts" ("Id") ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS "IX_PartBlankMaps_CanonicalBlankId" ON "PartBlankMaps" ("CanonicalBlankId");
CREATE INDEX IF NOT EXISTS "IX_PartBlankMaps_PartId" ON "PartBlankMaps" ("PartId");
CREATE UNIQUE INDEX IF NOT EXISTS "IX_PartBlankMaps_PartId_IsPrimary_IsActive" ON "PartBlankMaps" ("PartId", "IsPrimary", "IsActive") WHERE "IsPrimary" = 1 AND "IsActive" = 1;

CREATE TABLE IF NOT EXISTS "DemandBatches" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_DemandBatches" PRIMARY KEY AUTOINCREMENT,
    "Name" TEXT NOT NULL,
    "SourceFile" TEXT NULL,
    "FileHashSha256" TEXT NULL,
    "ImportedAt" TEXT NOT NULL,
    "PeriodFrom" TEXT NULL,
    "PeriodTo" TEXT NULL
);
CREATE TABLE IF NOT EXISTS "DemandItems" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_DemandItems" PRIMARY KEY AUTOINCREMENT,
    "DemandBatchId" INTEGER NOT NULL,
    "PartId" INTEGER NULL,
    "Ips" TEXT NOT NULL,
    "SourcePartName" TEXT NULL,
    "Quantity" TEXT NOT NULL,
    "DemandDate" TEXT NULL,
    "ProductionSystem" TEXT NULL,
    CONSTRAINT "FK_DemandItems_DemandBatches_DemandBatchId" FOREIGN KEY ("DemandBatchId") REFERENCES "DemandBatches" ("Id") ON DELETE CASCADE,
    CONSTRAINT "FK_DemandItems_Parts_PartId" FOREIGN KEY ("PartId") REFERENCES "Parts" ("Id")
);
CREATE INDEX IF NOT EXISTS "IX_DemandItems_DemandBatchId" ON "DemandItems" ("DemandBatchId");
CREATE INDEX IF NOT EXISTS "IX_DemandItems_DemandDate" ON "DemandItems" ("DemandDate");
CREATE INDEX IF NOT EXISTS "IX_DemandItems_Ips" ON "DemandItems" ("Ips");
CREATE INDEX IF NOT EXISTS "IX_DemandItems_PartId" ON "DemandItems" ("PartId");

CREATE TABLE IF NOT EXISTS "StockSnapshots" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_StockSnapshots" PRIMARY KEY AUTOINCREMENT,
    "SnapshotDate" TEXT NOT NULL,
    "SourceFile" TEXT NULL,
    "ImportedAt" TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS "StockItems" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_StockItems" PRIMARY KEY AUTOINCREMENT,
    "StockSnapshotId" INTEGER NOT NULL,
    "BlankAliasId" INTEGER NULL,
    "OneCCode" TEXT NOT NULL,
    "SourceName" TEXT NOT NULL,
    "Quantity" TEXT NOT NULL,
    "Unit" INTEGER NOT NULL,
    "Warehouse" TEXT NULL,
    CONSTRAINT "FK_StockItems_BlankAliases_BlankAliasId" FOREIGN KEY ("BlankAliasId") REFERENCES "BlankAliases" ("Id"),
    CONSTRAINT "FK_StockItems_StockSnapshots_StockSnapshotId" FOREIGN KEY ("StockSnapshotId") REFERENCES "StockSnapshots" ("Id") ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS "IX_StockItems_BlankAliasId" ON "StockItems" ("BlankAliasId");
CREATE INDEX IF NOT EXISTS "IX_StockItems_OneCCode" ON "StockItems" ("OneCCode");
CREATE INDEX IF NOT EXISTS "IX_StockItems_StockSnapshotId" ON "StockItems" ("StockSnapshotId");

CREATE TABLE IF NOT EXISTS "CalculationRuns" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_CalculationRuns" PRIMARY KEY AUTOINCREMENT,
    "StartedAt" TEXT NOT NULL,
    "FinishedAt" TEXT NULL,
    "DemandBatchId" INTEGER NOT NULL,
    "StockSnapshotId" INTEGER NULL,
    "Comment" TEXT NULL
);
CREATE TABLE IF NOT EXISTS "CalculationItems" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_CalculationItems" PRIMARY KEY AUTOINCREMENT,
    "CalculationRunId" INTEGER NOT NULL,
    "CanonicalBlankId" INTEGER NULL,
    "CanonicalName" TEXT NOT NULL,
    "PrimaryOneCCode" TEXT NULL,
    "OneCCodes" TEXT NOT NULL,
    "Unit" INTEGER NOT NULL,
    "TotalRequired" TEXT NOT NULL,
    "TotalStock" TEXT NOT NULL,
    "PurchaseQuantity" TEXT NOT NULL,
    "Status" INTEGER NOT NULL,
    "Comment" TEXT NULL,
    CONSTRAINT "FK_CalculationItems_CalculationRuns_CalculationRunId" FOREIGN KEY ("CalculationRunId") REFERENCES "CalculationRuns" ("Id") ON DELETE CASCADE,
    CONSTRAINT "FK_CalculationItems_CanonicalBlanks_CanonicalBlankId" FOREIGN KEY ("CanonicalBlankId") REFERENCES "CanonicalBlanks" ("Id")
);
CREATE TABLE IF NOT EXISTS "CalculationItemSources" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_CalculationItemSources" PRIMARY KEY AUTOINCREMENT,
    "CalculationItemId" INTEGER NOT NULL,
    "Ips" TEXT NOT NULL,
    "PartName" TEXT NOT NULL,
    "DemandQuantity" TEXT NOT NULL,
    "RequiredQuantity" TEXT NOT NULL,
    "Unit" INTEGER NOT NULL,
    CONSTRAINT "FK_CalculationItemSources_CalculationItems_CalculationItemId" FOREIGN KEY ("CalculationItemId") REFERENCES "CalculationItems" ("Id") ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS "IX_CalculationItems_CalculationRunId" ON "CalculationItems" ("CalculationRunId");
CREATE INDEX IF NOT EXISTS "IX_CalculationItems_CanonicalBlankId" ON "CalculationItems" ("CanonicalBlankId");
CREATE INDEX IF NOT EXISTS "IX_CalculationItemSources_CalculationItemId" ON "CalculationItemSources" ("CalculationItemId");

CREATE TABLE IF NOT EXISTS "I012CatalogEntries" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_I012CatalogEntries" PRIMARY KEY AUTOINCREMENT,
    "BlankType" INTEGER NOT NULL,
    "SizeKey" TEXT NOT NULL,
    "MaterialKey" TEXT NOT NULL,
    "Section" TEXT NOT NULL,
    "Source" TEXT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_I012CatalogEntries_BlankType_SizeKey_MaterialKey" ON "I012CatalogEntries" ("BlankType", "SizeKey", "MaterialKey");

CREATE TABLE IF NOT EXISTS "MskRecords" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_MskRecords" PRIMARY KEY AUTOINCREMENT,
    "Ips" TEXT NOT NULL,
    "Designation" TEXT NULL,
    "Name" TEXT NOT NULL,
    "FileName" TEXT NOT NULL,
    "ImportedAt" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_MskRecords_Ips" ON "MskRecords" ("Ips");

CREATE TABLE IF NOT EXISTS "ImportProfiles" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ImportProfiles" PRIMARY KEY AUTOINCREMENT,
    "ProfileName" TEXT NOT NULL,
    "ProfileType" TEXT NOT NULL,
    "SheetNamePattern" TEXT NULL,
    "ColumnMappingsJson" TEXT NOT NULL,
    "UpdatedAt" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_ImportProfiles_ProfileType_ProfileName" ON "ImportProfiles" ("ProfileType", "ProfileName");

CREATE TABLE IF NOT EXISTS "ImportRuns" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ImportRuns" PRIMARY KEY AUTOINCREMENT,
    "ImportType" TEXT NOT NULL,
    "SourceFile" TEXT NOT NULL,
    "FileHashSha256" TEXT NULL,
    "Status" INTEGER NOT NULL,
    "ReadRows" INTEGER NOT NULL,
    "AddedRows" INTEGER NOT NULL,
    "UpdatedRows" INTEGER NOT NULL,
    "SkippedRows" INTEGER NOT NULL,
    "ErrorRows" INTEGER NOT NULL,
    "StartedAt" TEXT NOT NULL,
    "FinishedAt" TEXT NULL
);
CREATE TABLE IF NOT EXISTS "ImportErrors" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ImportErrors" PRIMARY KEY AUTOINCREMENT,
    "ImportRunId" INTEGER NOT NULL,
    "RowNumber" INTEGER NOT NULL,
    "FieldName" TEXT NOT NULL,
    "Message" TEXT NOT NULL,
    CONSTRAINT "FK_ImportErrors_ImportRuns_ImportRunId" FOREIGN KEY ("ImportRunId") REFERENCES "ImportRuns" ("Id") ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS "IX_ImportErrors_ImportRunId" ON "ImportErrors" ("ImportRunId");

CREATE TABLE IF NOT EXISTS "ChangeHistory" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_ChangeHistory" PRIMARY KEY AUTOINCREMENT,
    "EntityType" TEXT NOT NULL,
    "EntityId" INTEGER NOT NULL,
    "FieldName" TEXT NOT NULL,
    "OldValue" TEXT NULL,
    "NewValue" TEXT NULL,
    "ChangedAt" TEXT NOT NULL,
    "Source" TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS "Settings" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_Settings" PRIMARY KEY AUTOINCREMENT,
    "Key" TEXT NOT NULL,
    "Value" TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS "IX_Settings_Key" ON "Settings" ("Key");
CREATE TABLE IF NOT EXISTS "Logs" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_Logs" PRIMARY KEY AUTOINCREMENT,
    "CreatedAt" TEXT NOT NULL,
    "Level" TEXT NOT NULL,
    "Message" TEXT NOT NULL,
    "Exception" TEXT NULL
);
""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
DROP TABLE IF EXISTS "Logs";
DROP TABLE IF EXISTS "Settings";
DROP TABLE IF EXISTS "ChangeHistory";
DROP TABLE IF EXISTS "ImportErrors";
DROP TABLE IF EXISTS "ImportRuns";
DROP TABLE IF EXISTS "ImportProfiles";
DROP TABLE IF EXISTS "MskRecords";
DROP TABLE IF EXISTS "I012CatalogEntries";
DROP TABLE IF EXISTS "CalculationItemSources";
DROP TABLE IF EXISTS "CalculationItems";
DROP TABLE IF EXISTS "CalculationRuns";
DROP TABLE IF EXISTS "StockItems";
DROP TABLE IF EXISTS "StockSnapshots";
DROP TABLE IF EXISTS "DemandItems";
DROP TABLE IF EXISTS "DemandBatches";
DROP TABLE IF EXISTS "PartBlankMaps";
DROP TABLE IF EXISTS "BlankAliases";
DROP TABLE IF EXISTS "CanonicalBlanks";
DROP TABLE IF EXISTS "Parts";
""");
    }
}
