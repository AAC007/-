using System.Windows;
using System.Globalization;
using System.IO;
using BlankDemandPlanner.Data;
using BlankDemandPlanner.Infrastructure;
using BlankDemandPlanner.Services;
using BlankDemandPlanner.UI.Services;
using BlankDemandPlanner.UI.ViewModels;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using FormsApplication = System.Windows.Forms.Application;
using UnhandledExceptionMode = System.Windows.Forms.UnhandledExceptionMode;

namespace BlankDemandPlanner.UI;

public partial class App : Application
{
    private IHost? _host;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        ConfigurePdfiumNativeSearchPath();
        var russianCulture = CultureInfo.GetCultureInfo("ru-RU");
        CultureInfo.DefaultThreadCurrentCulture = russianCulture;
        CultureInfo.DefaultThreadCurrentUICulture = russianCulture;

        DispatcherUnhandledException += (_, args) =>
        {
            Log.Error(args.Exception, "Unhandled dispatcher exception");
            MessageBox.Show("Произошла ошибка. Подробности записаны в лог.", "Планирование ЦМО", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            Log.Error(args.Exception, "Unobserved task exception");
            args.SetObserved();
        };
        FormsApplication.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        FormsApplication.ThreadException += (_, args) =>
        {
            Log.Error(args.Exception, "Unhandled Windows Forms exception");
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) => Log.Error(args.ExceptionObject as Exception, "Unhandled domain exception");

        var appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BlankDemandPlanner");
        Directory.CreateDirectory(appData);
        Directory.CreateDirectory(Path.Combine(appData, "Logs"));
        var legacyDatabasePath = Path.Combine(appData, "BlankDemandPlanner.db");
        var stableDataPath = FindStableDataDirectory(appData);
        Directory.CreateDirectory(stableDataPath);
        var databasePath = Path.Combine(stableDataPath, "BlankDemandPlanner.db");
        if (!File.Exists(databasePath) && File.Exists(legacyDatabasePath))
        {
            File.Copy(legacyDatabasePath, databasePath);
        }

        var backupPath = Path.Combine(appData, "Backups");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(Path.Combine(appData, "Logs", ".log"), rollingInterval: RollingInterval.Day, retainedFileCountLimit: 30)
            .CreateLogger();

        _host = Host.CreateDefaultBuilder(e.Args)
            .UseSerilog()
            .ConfigureServices(services =>
            {
                services.AddBlankDemandPlannerData(databasePath);
                services.AddBlankDemandPlannerServices();
                services.AddBlankDemandPlannerInfrastructure(databasePath, backupPath, keepBackups: 14);
                services.AddSingleton<IFileDialogService, FileDialogService>();
                services.AddSingleton<IAppAuthService, AppAuthService>();
                services.AddTransient<LoginWindow>();
                services.AddTransient<MainViewModel>();
                services.AddTransient<MainWindow>();
            })
            .Build();

        await _host.StartAsync();
        await using (var scope = _host.Services.CreateAsyncScope())
        {
            var db = scope.ServiceProvider.GetRequiredService<BlankDemandPlannerDbContext>();
            await db.Database.MigrateAsync();
            await db.Database.ExecuteSqlRawAsync("PRAGMA journal_mode=WAL;");
            await db.Database.ExecuteSqlRawAsync("PRAGMA busy_timeout=5000;");
        }

        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        var loginWindow = _host.Services.GetRequiredService<LoginWindow>();
        if (loginWindow.ShowDialog() != true)
        {
            Shutdown();
            return;
        }

        var window = _host.Services.GetRequiredService<MainWindow>();
        MainWindow = window;
        ShutdownMode = ShutdownMode.OnMainWindowClose;
        window.Show();
    }

    private static void ConfigurePdfiumNativeSearchPath()
    {
        var appDirectory = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var pdfiumDirectory = Path.Combine(appDirectory, "x64");
        var currentPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        var paths = currentPath.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);
        var additions = new[] { appDirectory, pdfiumDirectory }
            .Where(Directory.Exists)
            .Where(path => !paths.Any(existing => string.Equals(
                existing.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                path,
                StringComparison.OrdinalIgnoreCase)));

        Environment.SetEnvironmentVariable("PATH", string.Join(Path.PathSeparator, additions.Concat(paths)));
    }

    private static string FindStableDataDirectory(string fallbackAppData)
    {
        foreach (var basePath in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            var directory = new DirectoryInfo(basePath);
            for (var i = 0; directory is not null && i < 8; i++, directory = directory.Parent)
            {
                var candidate = Path.Combine(directory.FullName, "Данные для работы");
                if (Directory.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        return Path.Combine(fallbackAppData, "UserData");
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        if (_host is not null)
        {
            await _host.StopAsync(TimeSpan.FromSeconds(5));
            _host.Dispose();
        }

        Log.CloseAndFlush();
        base.OnExit(e);
    }
}
