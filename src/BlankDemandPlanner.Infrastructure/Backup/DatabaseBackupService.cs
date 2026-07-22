using BlankDemandPlanner.Core.Interfaces;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Infrastructure.Backup;

public sealed class DatabaseBackupService(string databasePath, string backupDirectory, int keepCount, ILogger<DatabaseBackupService> logger) : IDatabaseBackupService
{
    public async Task<string> BackupAsync(string reason, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(backupDirectory);
        var fileName = $"BlankDemandPlanner_{DateTime.Now:yyyy-MM-dd_HHmmss}.db";
        var destination = Path.Combine(backupDirectory, fileName);
        await using var source = File.Open(databasePath, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite);
        await using var target = File.Create(destination);
        await source.CopyToAsync(target, cancellationToken);
        logger.LogInformation("Database backup created: {Path}. Reason: {Reason}", destination, reason);
        await CleanupAsync(cancellationToken);
        return destination;
    }

    public Task CleanupAsync(CancellationToken cancellationToken)
    {
        if (!Directory.Exists(backupDirectory))
        {
            return Task.CompletedTask;
        }

        var files = Directory.GetFiles(backupDirectory, "BlankDemandPlanner_*.db")
            .OrderByDescending(File.GetCreationTimeUtc)
            .Skip(Math.Max(keepCount, 1))
            .ToList();

        foreach (var file in files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            File.Delete(file);
        }

        return Task.CompletedTask;
    }
}
