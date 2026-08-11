using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260804010000_AddPartExternalServiceFlags")]
public partial class AddPartExternalServiceFlags : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""ALTER TABLE "Parts" ADD COLUMN "RequiresNitriding" INTEGER NOT NULL DEFAULT 0;""");
        migrationBuilder.Sql("""ALTER TABLE "Parts" ADD COLUMN "RequiresHeatTreatment" INTEGER NOT NULL DEFAULT 0;""");
        migrationBuilder.Sql("""ALTER TABLE "Parts" ADD COLUMN "RequiresChemicalOxidation" INTEGER NOT NULL DEFAULT 0;""");
        migrationBuilder.Sql("""ALTER TABLE "Parts" ADD COLUMN "RequiresKeyway" INTEGER NOT NULL DEFAULT 0;""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
    }
}
