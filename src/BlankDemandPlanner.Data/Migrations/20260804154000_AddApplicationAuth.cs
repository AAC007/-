using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Infrastructure;

#nullable disable

namespace BlankDemandPlanner.Data.Migrations;

[DbContext(typeof(BlankDemandPlannerDbContext))]
[Migration("20260804154000_AddApplicationAuth")]
public partial class AddApplicationAuth : Migration
{
    protected override void Up(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.CreateTable(
            name: "AppUsers",
            columns: table => new
            {
                Id = table.Column<long>(type: "INTEGER", nullable: false)
                    .Annotation("Sqlite:Autoincrement", true),
                UserName = table.Column<string>(type: "TEXT", maxLength: 64, nullable: false),
                NormalizedUserName = table.Column<string>(type: "TEXT", maxLength: 64, nullable: false),
                DisplayName = table.Column<string>(type: "TEXT", maxLength: 256, nullable: false),
                PasswordHash = table.Column<string>(type: "TEXT", maxLength: 256, nullable: false),
                IsAdmin = table.Column<bool>(type: "INTEGER", nullable: false),
                IsActive = table.Column<bool>(type: "INTEGER", nullable: false),
                MustChangePassword = table.Column<bool>(type: "INTEGER", nullable: false),
                CreatedAt = table.Column<DateTime>(type: "TEXT", nullable: false),
                UpdatedAt = table.Column<DateTime>(type: "TEXT", nullable: false),
                LastLoginAt = table.Column<DateTime>(type: "TEXT", nullable: true)
            },
            constraints: table =>
            {
                table.PrimaryKey("PK_AppUsers", x => x.Id);
            });

        migrationBuilder.CreateTable(
            name: "AuthLoginAttempts",
            columns: table => new
            {
                Id = table.Column<long>(type: "INTEGER", nullable: false)
                    .Annotation("Sqlite:Autoincrement", true),
                UserName = table.Column<string>(type: "TEXT", maxLength: 64, nullable: false),
                NormalizedUserName = table.Column<string>(type: "TEXT", maxLength: 64, nullable: false),
                IsSuccess = table.Column<bool>(type: "INTEGER", nullable: false),
                FailureReason = table.Column<string>(type: "TEXT", maxLength: 256, nullable: true),
                MachineName = table.Column<string>(type: "TEXT", maxLength: 128, nullable: true),
                AttemptedAt = table.Column<DateTime>(type: "TEXT", nullable: false)
            },
            constraints: table =>
            {
                table.PrimaryKey("PK_AuthLoginAttempts", x => x.Id);
            });

        migrationBuilder.CreateTable(
            name: "AppUserPermissions",
            columns: table => new
            {
                Id = table.Column<long>(type: "INTEGER", nullable: false)
                    .Annotation("Sqlite:Autoincrement", true),
                AppUserId = table.Column<long>(type: "INTEGER", nullable: false),
                PageKey = table.Column<string>(type: "TEXT", maxLength: 64, nullable: false),
                CanRead = table.Column<bool>(type: "INTEGER", nullable: false),
                CanEdit = table.Column<bool>(type: "INTEGER", nullable: false)
            },
            constraints: table =>
            {
                table.PrimaryKey("PK_AppUserPermissions", x => x.Id);
                table.ForeignKey(
                    name: "FK_AppUserPermissions_AppUsers_AppUserId",
                    column: x => x.AppUserId,
                    principalTable: "AppUsers",
                    principalColumn: "Id",
                    onDelete: ReferentialAction.Cascade);
            });

        migrationBuilder.CreateIndex(
            name: "IX_AppUsers_NormalizedUserName",
            table: "AppUsers",
            column: "NormalizedUserName",
            unique: true);

        migrationBuilder.CreateIndex(
            name: "IX_AuthLoginAttempts_NormalizedUserName_AttemptedAt",
            table: "AuthLoginAttempts",
            columns: new[] { "NormalizedUserName", "AttemptedAt" });

        migrationBuilder.CreateIndex(
            name: "IX_AppUserPermissions_AppUserId_PageKey",
            table: "AppUserPermissions",
            columns: new[] { "AppUserId", "PageKey" },
            unique: true);
    }

    protected override void Down(MigrationBuilder migrationBuilder)
    {
        migrationBuilder.DropTable(name: "AppUserPermissions");
        migrationBuilder.DropTable(name: "AuthLoginAttempts");
        migrationBuilder.DropTable(name: "AppUsers");
    }
}
