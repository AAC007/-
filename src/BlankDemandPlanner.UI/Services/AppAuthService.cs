using System.Security.Cryptography;
using System.Text;
using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using BCryptNet = BCrypt.Net.BCrypt;

namespace BlankDemandPlanner.UI.Services;

public interface IAppAuthService
{
    AuthenticatedUser? CurrentUser { get; }
    IReadOnlyList<AppPagePermissionDefinition> PageDefinitions { get; }
    Task<bool> HasAnyUsersAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<string>> GetLoginUserNamesAsync(CancellationToken cancellationToken = default);
    Task<string?> GetLastLoginUserNameAsync(CancellationToken cancellationToken = default);
    Task<AuthResult> SetupFirstAdminAsync(string userName, string displayName, string password, CancellationToken cancellationToken = default);
    Task<AuthResult> LoginAsync(string userName, string password, CancellationToken cancellationToken = default);
    bool CanRead(string pageKey);
    bool CanEdit(string pageKey);
    Task<IReadOnlyList<AppUserAdminRow>> GetUsersAsync(CancellationToken cancellationToken = default);
    Task<AuthResult> SaveUserAsync(AppUserEditRequest request, CancellationToken cancellationToken = default);
    Task<AuthResult> DeleteUserAsync(long userId, CancellationToken cancellationToken = default);
    Task<AuthResult> ResetPasswordAsync(long userId, string newPassword, CancellationToken cancellationToken = default);
    Task<AuthResult> ChangeOwnPasswordAsync(string currentPassword, string newPassword, CancellationToken cancellationToken = default);
    string GeneratePassword(int length = 14);
}

public sealed class AppAuthService(IDbContextFactory<BlankDemandPlannerDbContext> dbContextFactory, ILogger<AppAuthService> logger) : IAppAuthService
{
    private const int BcryptWorkFactor = 12;
    private const int MaxFailedAttempts = 5;
    private static readonly TimeSpan LockoutWindow = TimeSpan.FromMinutes(15);

    public AuthenticatedUser? CurrentUser { get; private set; }

    public IReadOnlyList<AppPagePermissionDefinition> PageDefinitions { get; } =
    [
        new("Dashboard", "Главная"),
        new("Demand", "Потребность"),
        new("Calculation", "Расчет материалов"),
        new("Library", "Библиотека"),
        new("Normalization", "НСИ"),
        new("BlankSelection", "Подбор заготовок"),
        new("Planning", "Планирование"),
        new("WorkshopReport", "Отчет"),
        new("Plan", "План"),
        new("History", "История"),
        new("Settings", "Настройки"),
        new("DeveloperMode", "Режим разработчика")
    ];

    public Task<bool> HasAnyUsersAsync(CancellationToken cancellationToken = default) =>
        HasAnyUsersCoreAsync(cancellationToken);

    public async Task<IReadOnlyList<string>> GetLoginUserNamesAsync(CancellationToken cancellationToken = default)
    {
        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        return await dbContext.AppUsers
            .Where(x => x.IsActive)
            .OrderByDescending(x => x.LastLoginAt)
            .ThenBy(x => x.UserName)
            .Select(x => x.UserName)
            .ToListAsync(cancellationToken);
    }

    public async Task<string?> GetLastLoginUserNameAsync(CancellationToken cancellationToken = default)
    {
        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        return await dbContext.Settings
            .Where(x => x.Key == LastLoginUserNameSettingKey)
            .Select(x => x.Value)
            .FirstOrDefaultAsync(cancellationToken);
    }

    public async Task<AuthResult> SetupFirstAdminAsync(string userName, string displayName, string password, CancellationToken cancellationToken = default)
    {
        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        if (await dbContext.AppUsers.AnyAsync(cancellationToken))
        {
            return AuthResult.Fail("Первичный администратор уже создан.");
        }

        var validation = ValidateUserNameAndPassword(userName, password);
        if (!validation.IsSuccess)
        {
            return validation;
        }

        var user = new AppUser
        {
            UserName = userName.Trim(),
            NormalizedUserName = NormalizeUserName(userName),
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? userName.Trim() : displayName.Trim(),
            PasswordHash = HashPassword(password),
            IsAdmin = true,
            IsActive = true,
            MustChangePassword = false,
            CreatedAt = DateTime.UtcNow,
            UpdatedAt = DateTime.UtcNow
        };
        dbContext.AppUsers.Add(user);
        await dbContext.SaveChangesAsync(cancellationToken);
        return await LoginAsync(userName, password, cancellationToken);
    }

    public async Task<AuthResult> LoginAsync(string userName, string password, CancellationToken cancellationToken = default)
    {
        var normalized = NormalizeUserName(userName);
        if (string.IsNullOrWhiteSpace(normalized) || string.IsNullOrWhiteSpace(password))
        {
            await LogAttemptAsync(userName, false, "Пустой логин или пароль", cancellationToken);
            return AuthResult.Fail("Введите логин и пароль.");
        }

        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        var lockoutStartedAt = DateTime.UtcNow.Subtract(LockoutWindow);
        var recentFailures = await dbContext.AuthLoginAttempts
            .CountAsync(x => x.NormalizedUserName == normalized &&
                !x.IsSuccess &&
                x.AttemptedAt >= lockoutStartedAt, cancellationToken);
        if (recentFailures >= MaxFailedAttempts)
        {
            await LogAttemptAsync(userName, false, "Rate limit", cancellationToken);
            return AuthResult.Fail("Слишком много неудачных попыток. Повторите вход через 15 минут.");
        }

        var user = await dbContext.AppUsers
            .Include(x => x.Permissions)
            .FirstOrDefaultAsync(x => x.NormalizedUserName == normalized, cancellationToken);
        if (user is null || !user.IsActive)
        {
            await LogAttemptAsync(userName, false, "Пользователь не найден или отключен", cancellationToken);
            return AuthResult.Fail("Неверный логин или пароль.");
        }

        if (!BCryptNet.Verify(password, user.PasswordHash))
        {
            logger.LogWarning("Failed login for {UserName} from {Machine}", userName, Environment.MachineName);
            await LogAttemptAsync(userName, false, "Неверный пароль", cancellationToken);
            return AuthResult.Fail("Неверный логин или пароль.");
        }

        user.LastLoginAt = DateTime.UtcNow;
        await LogAttemptAsync(userName, true, null, cancellationToken);
        await SaveLastLoginUserNameAsync(dbContext, user.UserName, cancellationToken);
        await dbContext.SaveChangesAsync(cancellationToken);
        CurrentUser = AuthenticatedUser.FromEntity(user);
        return AuthResult.Ok("Вход выполнен.");
    }

    public bool CanRead(string pageKey) =>
        CurrentUser?.IsAdmin == true || CurrentUser?.Permissions.TryGetValue(pageKey, out var permission) == true && permission.CanRead;

    public bool CanEdit(string pageKey) =>
        CurrentUser?.IsAdmin == true || CurrentUser?.Permissions.TryGetValue(pageKey, out var permission) == true && permission.CanEdit;

    public async Task<IReadOnlyList<AppUserAdminRow>> GetUsersAsync(CancellationToken cancellationToken = default)
    {
        if (CurrentUser?.IsAdmin != true)
        {
            return [];
        }

        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        return await dbContext.AppUsers
            .Include(x => x.Permissions)
            .OrderByDescending(x => x.IsAdmin)
            .ThenBy(x => x.UserName)
            .Select(x => new AppUserAdminRow(
                x.Id,
                x.UserName,
                x.DisplayName,
                x.IsAdmin,
                x.IsActive,
                x.MustChangePassword,
                x.LastLoginAt,
                x.Permissions.Select(p => new AppUserPermissionRow(p.PageKey, p.CanRead, p.CanEdit)).ToList()))
            .ToListAsync(cancellationToken);
    }

    public async Task<AuthResult> SaveUserAsync(AppUserEditRequest request, CancellationToken cancellationToken = default)
    {
        if (CurrentUser?.IsAdmin != true)
        {
            return AuthResult.Fail("Недостаточно прав.");
        }

        if (string.IsNullOrWhiteSpace(request.UserName) || request.UserName.Trim().Length < 3)
        {
            return AuthResult.Fail("Логин должен быть не короче 3 символов.");
        }

        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        AppUser user;
        var normalized = NormalizeUserName(request.UserName);
        if (request.UserId is null or 0)
        {
            if (string.IsNullOrWhiteSpace(request.NewPassword))
            {
                return AuthResult.Fail("Для нового пользователя нужен пароль.");
            }

            var validation = ValidateUserNameAndPassword(request.UserName, request.NewPassword);
            if (!validation.IsSuccess)
            {
                return validation;
            }

            if (await dbContext.AppUsers.AnyAsync(x => x.NormalizedUserName == normalized, cancellationToken))
            {
                return AuthResult.Fail("Пользователь с таким логином уже существует.");
            }

            user = new AppUser
            {
                CreatedAt = DateTime.UtcNow,
                PasswordHash = HashPassword(request.NewPassword),
                MustChangePassword = true
            };
            dbContext.AppUsers.Add(user);
        }
        else
        {
            user = await dbContext.AppUsers
                .Include(x => x.Permissions)
                .FirstOrDefaultAsync(x => x.Id == request.UserId.Value, cancellationToken)
                ?? throw new InvalidOperationException("Пользователь не найден.");
            if (!string.Equals(user.NormalizedUserName, normalized, StringComparison.OrdinalIgnoreCase) &&
                await dbContext.AppUsers.AnyAsync(x => x.NormalizedUserName == normalized, cancellationToken))
            {
                return AuthResult.Fail("Пользователь с таким логином уже существует.");
            }
        }

        user.UserName = request.UserName.Trim();
        user.NormalizedUserName = normalized;
        user.DisplayName = string.IsNullOrWhiteSpace(request.DisplayName) ? request.UserName.Trim() : request.DisplayName.Trim();
        user.IsAdmin = request.IsAdmin;
        user.IsActive = request.IsActive;
        user.UpdatedAt = DateTime.UtcNow;
        ApplyPermissions(user, request.Permissions);
        await dbContext.SaveChangesAsync(cancellationToken);
        if (CurrentUser?.Id == user.Id)
        {
            CurrentUser = AuthenticatedUser.FromEntity(user);
        }

        return AuthResult.Ok("Пользователь сохранен.");
    }

    public async Task<AuthResult> DeleteUserAsync(long userId, CancellationToken cancellationToken = default)
    {
        if (CurrentUser?.IsAdmin != true)
        {
            return AuthResult.Fail("Недостаточно прав.");
        }

        if (CurrentUser.Id == userId)
        {
            return AuthResult.Fail("Нельзя удалить текущего пользователя.");
        }

        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        var user = await dbContext.AppUsers.FirstOrDefaultAsync(x => x.Id == userId, cancellationToken);
        if (user is null)
        {
            return AuthResult.Fail("Пользователь не найден.");
        }

        if (user.IsAdmin)
        {
            var activeAdminCount = await dbContext.AppUsers.CountAsync(x => x.IsActive && x.IsAdmin, cancellationToken);
            if (activeAdminCount <= 1)
            {
                return AuthResult.Fail("Нельзя удалить последнего активного администратора.");
            }
        }

        user.IsActive = false;
        user.UpdatedAt = DateTime.UtcNow;
        await dbContext.SaveChangesAsync(cancellationToken);
        return AuthResult.Ok("Пользователь удален из активного списка.");
    }

    public async Task<AuthResult> ResetPasswordAsync(long userId, string newPassword, CancellationToken cancellationToken = default)
    {
        if (CurrentUser?.IsAdmin != true)
        {
            return AuthResult.Fail("Недостаточно прав.");
        }

        var validation = ValidatePassword(newPassword);
        if (!validation.IsSuccess)
        {
            return validation;
        }

        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        var user = await dbContext.AppUsers.FirstOrDefaultAsync(x => x.Id == userId, cancellationToken);
        if (user is null)
        {
            return AuthResult.Fail("Пользователь не найден.");
        }

        user.PasswordHash = HashPassword(newPassword);
        user.MustChangePassword = true;
        user.UpdatedAt = DateTime.UtcNow;
        await dbContext.SaveChangesAsync(cancellationToken);
        return AuthResult.Ok("Пароль сброшен администратором.");
    }

    public async Task<AuthResult> ChangeOwnPasswordAsync(string currentPassword, string newPassword, CancellationToken cancellationToken = default)
    {
        if (CurrentUser is null)
        {
            return AuthResult.Fail("Вход не выполнен.");
        }

        var validation = ValidatePassword(newPassword);
        if (!validation.IsSuccess)
        {
            return validation;
        }

        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        var currentUser = CurrentUser;
        var user = await dbContext.AppUsers.FirstOrDefaultAsync(x => x.Id == currentUser.Id, cancellationToken);
        if (user is null || !BCryptNet.Verify(currentPassword, user.PasswordHash))
        {
            return AuthResult.Fail("Текущий пароль указан неверно.");
        }

        user.PasswordHash = HashPassword(newPassword);
        user.MustChangePassword = false;
        user.UpdatedAt = DateTime.UtcNow;
        await dbContext.SaveChangesAsync(cancellationToken);
        CurrentUser = AuthenticatedUser.FromEntity(user);
        return AuthResult.Ok("Пароль изменен.");
    }

    public string GeneratePassword(int length = 14)
    {
        const string alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%";
        Span<byte> bytes = stackalloc byte[length];
        RandomNumberGenerator.Fill(bytes);
        var builder = new StringBuilder(length);
        foreach (var value in bytes)
        {
            builder.Append(alphabet[value % alphabet.Length]);
        }

        return builder.ToString();
    }

    private static string HashPassword(string password) =>
        BCryptNet.HashPassword(password, workFactor: BcryptWorkFactor);

    private const string LastLoginUserNameSettingKey = "Auth.LastUserName";

    private static AuthResult ValidateUserNameAndPassword(string userName, string password) =>
        string.IsNullOrWhiteSpace(userName) || userName.Trim().Length < 3
            ? AuthResult.Fail("Логин должен быть не короче 3 символов.")
            : ValidatePassword(password);

    private static AuthResult ValidatePassword(string password)
    {
        if (string.IsNullOrWhiteSpace(password) || password.Length < 10)
        {
            return AuthResult.Fail("Пароль должен быть не короче 10 символов.");
        }

        if (!password.Any(char.IsUpper) || !password.Any(char.IsLower) || !password.Any(char.IsDigit))
        {
            return AuthResult.Fail("Пароль должен содержать строчные, заглавные буквы и цифры.");
        }

        return AuthResult.Ok(string.Empty);
    }

    private static string NormalizeUserName(string value) => value.Trim().ToUpperInvariant();

    private async Task LogAttemptAsync(string userName, bool success, string? reason, CancellationToken cancellationToken)
    {
        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        dbContext.AuthLoginAttempts.Add(new AuthLoginAttempt
        {
            UserName = userName.Trim(),
            NormalizedUserName = NormalizeUserName(userName),
            IsSuccess = success,
            FailureReason = reason,
            MachineName = Environment.MachineName,
            AttemptedAt = DateTime.UtcNow
        });
        await dbContext.SaveChangesAsync(cancellationToken);
    }

    private async Task<bool> HasAnyUsersCoreAsync(CancellationToken cancellationToken)
    {
        await using var dbContext = await dbContextFactory.CreateDbContextAsync(cancellationToken);
        return await dbContext.AppUsers.AnyAsync(cancellationToken);
    }

    private static async Task SaveLastLoginUserNameAsync(BlankDemandPlannerDbContext dbContext, string userName, CancellationToken cancellationToken)
    {
        var setting = await dbContext.Settings.FirstOrDefaultAsync(x => x.Key == LastLoginUserNameSettingKey, cancellationToken);
        if (setting is null)
        {
            dbContext.Settings.Add(new AppSetting { Key = LastLoginUserNameSettingKey, Value = userName });
            return;
        }

        setting.Value = userName;
    }

    private void ApplyPermissions(AppUser user, IReadOnlyCollection<AppUserPermissionRow> permissions)
    {
        if (user.IsAdmin)
        {
            user.Permissions.Clear();
            return;
        }

        foreach (var definition in PageDefinitions)
        {
            var source = permissions.FirstOrDefault(x => x.PageKey == definition.PageKey);
            var target = user.Permissions.FirstOrDefault(x => x.PageKey == definition.PageKey);
            if (target is null)
            {
                target = new AppUserPermission { PageKey = definition.PageKey, AppUser = user };
                user.Permissions.Add(target);
            }

            target.CanRead = source?.CanRead == true || definition.PageKey == "Settings";
            target.CanEdit = source?.CanEdit == true && target.CanRead;
        }
    }
}

public sealed record AuthenticatedUser(long Id, string UserName, string DisplayName, bool IsAdmin, bool MustChangePassword, IReadOnlyDictionary<string, AppUserPermissionRow> Permissions)
{
    public static AuthenticatedUser FromEntity(AppUser user) =>
        new(
            user.Id,
            user.UserName,
            user.DisplayName,
            user.IsAdmin,
            user.MustChangePassword,
            user.Permissions.ToDictionary(x => x.PageKey, x => new AppUserPermissionRow(x.PageKey, x.CanRead, x.CanEdit)));
}

public sealed record AuthResult(bool IsSuccess, string Message)
{
    public static AuthResult Ok(string message) => new(true, message);
    public static AuthResult Fail(string message) => new(false, message);
}

public sealed record AppPagePermissionDefinition(string PageKey, string Title);

public sealed record AppUserPermissionRow(string PageKey, bool CanRead, bool CanEdit);

public sealed record AppUserAdminRow(
    long Id,
    string UserName,
    string DisplayName,
    bool IsAdmin,
    bool IsActive,
    bool MustChangePassword,
    DateTime? LastLoginAt,
    IReadOnlyList<AppUserPermissionRow> Permissions);

public sealed record AppUserEditRequest(
    long? UserId,
    string UserName,
    string DisplayName,
    bool IsAdmin,
    bool IsActive,
    string? NewPassword,
    IReadOnlyCollection<AppUserPermissionRow> Permissions);
