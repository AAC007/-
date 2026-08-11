using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260728182500_AddOneCPriceItems")]
public partial class AddOneCPriceItems : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
CREATE TABLE IF NOT EXISTS "OneCPriceItems" (
    "Id" INTEGER NOT NULL CONSTRAINT "PK_OneCPriceItems" PRIMARY KEY AUTOINCREMENT,
    "LookupKey" TEXT NOT NULL,
    "Code" TEXT NOT NULL,
    "Article" TEXT NOT NULL,
    "Price" TEXT NOT NULL,
    "Currency" TEXT NOT NULL,
    "PriceType" TEXT NOT NULL,
    "SyncedAt" TEXT NOT NULL
);
""");
        migrationBuilder.Sql("""
CREATE UNIQUE INDEX IF NOT EXISTS "IX_OneCPriceItems_LookupKey" ON "OneCPriceItems" ("LookupKey");
""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""DROP TABLE IF EXISTS "OneCPriceItems";""");
    }
}
