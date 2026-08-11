using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260810120000_AddPartBlankSupplyRequirements")]
public partial class AddPartBlankSupplyRequirements : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""ALTER TABLE "Parts" ADD COLUMN "BlankSupplyRequiresHeatTreatment" INTEGER NOT NULL DEFAULT 0;""");
        migrationBuilder.Sql("""ALTER TABLE "Parts" ADD COLUMN "BlankSupplyRequiresLaserCutting" INTEGER NOT NULL DEFAULT 0;""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
    }
}
