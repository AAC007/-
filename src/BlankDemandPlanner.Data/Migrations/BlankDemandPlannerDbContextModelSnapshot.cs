using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
public sealed class BlankDemandPlannerDbContextModelSnapshot : ModelSnapshot
{
    protected override void BuildModel(ModelBuilder modelBuilder)
    {
        modelBuilder.HasAnnotation("ProductVersion", "8.0.7");
    }
}
