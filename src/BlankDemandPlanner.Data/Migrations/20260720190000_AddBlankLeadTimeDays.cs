using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260720190000_AddBlankLeadTimeDays")]
public partial class AddBlankLeadTimeDays : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.Sql("""
ALTER TABLE "PartBlankMaps" ADD COLUMN "BlankLeadTimeDays" INTEGER NOT NULL DEFAULT 30;
""");
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        // SQLite cannot drop columns reliably across deployed versions without rebuilding the table.
    }
}
