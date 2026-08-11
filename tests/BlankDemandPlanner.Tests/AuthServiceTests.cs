using BlankDemandPlanner.Data;
using BlankDemandPlanner.UI.Services;
using FluentAssertions;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.Abstractions;

namespace BlankDemandPlanner.Tests;

public sealed class AuthServiceTests
{
    [Fact]
    public async Task First_admin_password_is_hashed_and_login_uses_bcrypt()
    {
        await using var factory = new TestDbContextFactory();
        await using (var db = await factory.CreateDbContextAsync())
        {
            await db.Database.EnsureCreatedAsync();
        }

        var service = new AppAuthService(factory, NullLogger<AppAuthService>.Instance);
        var result = await service.SetupFirstAdminAsync("admin", "Администратор", "AdminPass123!");

        result.IsSuccess.Should().BeTrue();
        await using var checkDb = await factory.CreateDbContextAsync();
        var user = await checkDb.AppUsers.SingleAsync();
        user.PasswordHash.Should().NotBe("AdminPass123!");
        user.PasswordHash.Should().StartWith("$2");
        service.CurrentUser.Should().NotBeNull();
        service.CurrentUser!.IsAdmin.Should().BeTrue();
    }

    [Fact]
    public async Task Login_is_rate_limited_after_failed_attempts()
    {
        await using var factory = new TestDbContextFactory();
        await using (var db = await factory.CreateDbContextAsync())
        {
            await db.Database.EnsureCreatedAsync();
        }

        var service = new AppAuthService(factory, NullLogger<AppAuthService>.Instance);
        (await service.SetupFirstAdminAsync("admin", "Администратор", "AdminPass123!")).IsSuccess.Should().BeTrue();

        for (var i = 0; i < 5; i++)
        {
            (await service.LoginAsync("admin", "WrongPass123!")).IsSuccess.Should().BeFalse();
        }

        var blocked = await service.LoginAsync("admin", "AdminPass123!");
        blocked.IsSuccess.Should().BeFalse();
        blocked.Message.Should().Contain("Слишком много");
    }

    [Fact]
    public async Task First_admin_setup_logs_in_with_sqlite_provider()
    {
        await using var factory = new SqliteTestDbContextFactory();
        await using (var db = await factory.CreateDbContextAsync())
        {
            await db.Database.MigrateAsync();
        }

        var service = new AppAuthService(factory, NullLogger<AppAuthService>.Instance);
        var result = await service.SetupFirstAdminAsync("admin", "Администратор", "AdminPass123!");

        result.IsSuccess.Should().BeTrue();
        service.CurrentUser.Should().NotBeNull();
        service.CurrentUser!.UserName.Should().Be("admin");
    }

    [Fact]
    public async Task Successful_login_remembers_last_user_name()
    {
        await using var factory = new TestDbContextFactory();
        await using (var db = await factory.CreateDbContextAsync())
        {
            await db.Database.EnsureCreatedAsync();
        }

        var service = new AppAuthService(factory, NullLogger<AppAuthService>.Instance);
        (await service.SetupFirstAdminAsync("admin", "Администратор", "AdminPass123!")).IsSuccess.Should().BeTrue();
        (await service.SaveUserAsync(new AppUserEditRequest(
            null,
            "planner",
            "Планировщик",
            false,
            true,
            "PlannerPass123!",
            [new AppUserPermissionRow("Settings", true, false)]))).IsSuccess.Should().BeTrue();

        (await service.LoginAsync("planner", "PlannerPass123!")).IsSuccess.Should().BeTrue();

        (await service.GetLastLoginUserNameAsync()).Should().Be("planner");
    }

    [Fact]
    public async Task Login_user_list_contains_only_active_users()
    {
        await using var factory = new TestDbContextFactory();
        await using (var db = await factory.CreateDbContextAsync())
        {
            await db.Database.EnsureCreatedAsync();
        }

        var service = new AppAuthService(factory, NullLogger<AppAuthService>.Instance);
        (await service.SetupFirstAdminAsync("admin", "Администратор", "AdminPass123!")).IsSuccess.Should().BeTrue();
        (await service.SaveUserAsync(new AppUserEditRequest(null, "active", "Активный", false, true, "ActivePass123!", []))).IsSuccess.Should().BeTrue();
        (await service.SaveUserAsync(new AppUserEditRequest(null, "blocked", "Отключенный", false, false, "BlockedPass123!", []))).IsSuccess.Should().BeTrue();

        var users = await service.GetLoginUserNamesAsync();

        users.Should().Contain("admin");
        users.Should().Contain("active");
        users.Should().NotContain("blocked");
    }

    [Fact]
    public async Task Admin_can_delete_user_by_deactivating_account()
    {
        await using var factory = new TestDbContextFactory();
        await using (var db = await factory.CreateDbContextAsync())
        {
            await db.Database.EnsureCreatedAsync();
        }

        var service = new AppAuthService(factory, NullLogger<AppAuthService>.Instance);
        (await service.SetupFirstAdminAsync("admin", "Администратор", "AdminPass123!")).IsSuccess.Should().BeTrue();
        (await service.SaveUserAsync(new AppUserEditRequest(null, "planner", "Планировщик", false, true, "PlannerPass123!", []))).IsSuccess.Should().BeTrue();
        var user = (await service.GetUsersAsync()).Single(x => x.UserName == "planner");

        var result = await service.DeleteUserAsync(user.Id);

        result.IsSuccess.Should().BeTrue();
        (await service.GetLoginUserNamesAsync()).Should().NotContain("planner");
        (await service.LoginAsync("planner", "PlannerPass123!")).IsSuccess.Should().BeFalse();
        (await service.DeleteUserAsync(service.CurrentUser!.Id)).IsSuccess.Should().BeFalse();
    }

    private sealed class TestDbContextFactory : IDbContextFactory<BlankDemandPlannerDbContext>, IAsyncDisposable
    {
        private readonly DbContextOptions<BlankDemandPlannerDbContext> _options;

        public TestDbContextFactory()
        {
            _options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseInMemoryDatabase(Guid.NewGuid().ToString("N"))
                .Options;
        }

        public BlankDemandPlannerDbContext CreateDbContext() => new(_options);

        public ValueTask<BlankDemandPlannerDbContext> CreateDbContextAsync(CancellationToken cancellationToken = default) =>
            ValueTask.FromResult(CreateDbContext());

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class SqliteTestDbContextFactory : IDbContextFactory<BlankDemandPlannerDbContext>, IAsyncDisposable
    {
        private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"blank_auth_{Guid.NewGuid():N}.db");
        private readonly DbContextOptions<BlankDemandPlannerDbContext> _options;

        public SqliteTestDbContextFactory()
        {
            _options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={_databasePath}")
                .Options;
        }

        public BlankDemandPlannerDbContext CreateDbContext() => new(_options);

        public ValueTask<BlankDemandPlannerDbContext> CreateDbContextAsync(CancellationToken cancellationToken = default) =>
            ValueTask.FromResult(CreateDbContext());

        public ValueTask DisposeAsync()
        {
            SqliteConnection.ClearAllPools();
            foreach (var path in new[] { _databasePath, $"{_databasePath}-shm", $"{_databasePath}-wal" })
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }

            return ValueTask.CompletedTask;
        }
    }
}
