using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260716150000_AddDemandDisplayFields")]
public partial class AddDemandDisplayFields : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
ALTER TABLE "DemandItems" ADD COLUMN "Project" TEXT NULL;
""");

        migrationBuilder.Sql("""
ALTER TABLE "DemandItems" ADD COLUMN "SerialNumber" TEXT NULL;
""");

        migrationBuilder.Sql("""
ALTER TABLE "DemandItems" ADD COLUMN "Unit" TEXT NULL;
""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        // SQLite cannot drop columns reliably across deployed versions without rebuilding the table.
    }
}
