using BlankDemandPlanner.Core.Entities;
using Microsoft.EntityFrameworkCore;

namespace BlankDemandPlanner.Data;

public sealed class BlankDemandPlannerDbContext(DbContextOptions<BlankDemandPlannerDbContext> options) : DbContext(options)
{
    public DbSet<Part> Parts => Set<Part>();
    public DbSet<CanonicalBlank> CanonicalBlanks => Set<CanonicalBlank>();
    public DbSet<BlankAlias> BlankAliases => Set<BlankAlias>();
    public DbSet<PartBlankMap> PartBlankMaps => Set<PartBlankMap>();
    public DbSet<DemandBatch> DemandBatches => Set<DemandBatch>();
    public DbSet<DemandItem> DemandItems => Set<DemandItem>();
    public DbSet<StockSnapshot> StockSnapshots => Set<StockSnapshot>();
    public DbSet<StockItem> StockItems => Set<StockItem>();
    public DbSet<OneCPriceItem> OneCPriceItems => Set<OneCPriceItem>();
    public DbSet<CalculationRun> CalculationRuns => Set<CalculationRun>();
    public DbSet<CalculationItem> CalculationItems => Set<CalculationItem>();
    public DbSet<CalculationItemSource> CalculationItemSources => Set<CalculationItemSource>();
    public DbSet<I012CatalogEntry> I012CatalogEntries => Set<I012CatalogEntry>();
    public DbSet<MskRecord> MskRecords => Set<MskRecord>();
    public DbSet<ImportProfile> ImportProfiles => Set<ImportProfile>();
    public DbSet<ImportRun> ImportRuns => Set<ImportRun>();
    public DbSet<ImportError> ImportErrors => Set<ImportError>();
    public DbSet<ChangeHistory> ChangeHistory => Set<ChangeHistory>();
    public DbSet<AppSetting> Settings => Set<AppSetting>();
    public DbSet<AppLog> Logs => Set<AppLog>();
    public DbSet<ProductionEquipment> ProductionEquipment => Set<ProductionEquipment>();
    public DbSet<ProductionEmployee> ProductionEmployees => Set<ProductionEmployee>();
    public DbSet<ProductionRouteOperation> ProductionRouteOperations => Set<ProductionRouteOperation>();
    public DbSet<ProductionScheduleEntry> ProductionScheduleEntries => Set<ProductionScheduleEntry>();
    public DbSet<AppUser> AppUsers => Set<AppUser>();
    public DbSet<AppUserPermission> AppUserPermissions => Set<AppUserPermission>();
    public DbSet<AuthLoginAttempt> AuthLoginAttempts => Set<AuthLoginAttempt>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Part>(entity =>
        {
            entity.Property(x => x.Ips).HasMaxLength(64);
            entity.Property(x => x.Name).HasMaxLength(512);
            entity.HasIndex(x => x.Ips).IsUnique();
        });

        modelBuilder.Entity<CanonicalBlank>(entity =>
        {
            entity.Property(x => x.CanonicalName).HasMaxLength(1024);
            entity.Property(x => x.CanonicalKey).HasMaxLength(512);
            entity.HasIndex(x => x.CanonicalKey).IsUnique();
        });

        modelBuilder.Entity<BlankAlias>(entity =>
        {
            entity.Property(x => x.OneCCode).HasMaxLength(64);
            entity.Property(x => x.SourceName).HasMaxLength(1024);
            entity.HasIndex(x => x.OneCCode).IsUnique();
            entity.HasIndex(x => x.CanonicalBlankId);
        });

        modelBuilder.Entity<PartBlankMap>(entity =>
        {
            entity.HasIndex(x => x.PartId);
            entity.HasIndex(x => x.CanonicalBlankId);
            entity.HasIndex(x => new { x.PartId, x.IsPrimary, x.IsActive })
                .IsUnique()
                .HasFilter("\"IsPrimary\" = 1 AND \"IsActive\" = 1");
        });

        modelBuilder.Entity<DemandItem>(entity =>
        {
            entity.Property(x => x.Project).HasMaxLength(512);
            entity.Property(x => x.SerialNumber).HasMaxLength(128);
            entity.Property(x => x.Unit).HasMaxLength(32);
            entity.Property(x => x.ProductionSystem).HasMaxLength(128);
            entity.HasIndex(x => x.Ips);
            entity.HasIndex(x => x.DemandDate);
        });

        modelBuilder.Entity<StockItem>(entity =>
        {
            entity.HasIndex(x => x.OneCCode);
            entity.HasIndex(x => x.BlankAliasId);
        });

        modelBuilder.Entity<OneCPriceItem>(entity =>
        {
            entity.Property(x => x.LookupKey).HasMaxLength(64);
            entity.Property(x => x.Code).HasMaxLength(64);
            entity.Property(x => x.Article).HasMaxLength(64);
            entity.Property(x => x.Currency).HasMaxLength(32);
            entity.Property(x => x.PriceType).HasMaxLength(128);
            entity.HasIndex(x => x.LookupKey).IsUnique();
        });

        modelBuilder.Entity<I012CatalogEntry>(entity =>
        {
            entity.HasIndex(x => new { x.BlankType, x.SizeKey, x.MaterialKey }).IsUnique();
        });

        modelBuilder.Entity<MskRecord>().HasIndex(x => x.Ips).IsUnique();
        modelBuilder.Entity<ImportProfile>().HasIndex(x => new { x.ProfileType, x.ProfileName }).IsUnique();
        modelBuilder.Entity<AppSetting>().HasIndex(x => x.Key).IsUnique();

        modelBuilder.Entity<ProductionEquipment>(entity =>
        {
            entity.Property(x => x.Code).HasMaxLength(64);
            entity.Property(x => x.Name).HasMaxLength(256);
            entity.Property(x => x.ResourceGroup).HasMaxLength(128);
            entity.HasIndex(x => x.Code).IsUnique();
            entity.HasIndex(x => x.ResourceGroup);
        });

        modelBuilder.Entity<ProductionEmployee>(entity =>
        {
            entity.Property(x => x.PersonnelNumber).HasMaxLength(64);
            entity.Property(x => x.FullName).HasMaxLength(256);
            entity.Property(x => x.Specialty).HasMaxLength(128);
            entity.HasIndex(x => x.PersonnelNumber).IsUnique();
            entity.HasIndex(x => x.Specialty);
        });

        modelBuilder.Entity<ProductionRouteOperation>(entity =>
        {
            entity.Property(x => x.Ips).HasMaxLength(64);
            entity.Property(x => x.OperationCode).HasMaxLength(64);
            entity.Property(x => x.EquipmentGroup).HasMaxLength(128);
            entity.Property(x => x.RequiredSpecialty).HasMaxLength(128);
            entity.HasIndex(x => new { x.Ips, x.Sequence }).IsUnique();
            entity.HasIndex(x => x.Ips);
        });

        modelBuilder.Entity<ProductionScheduleEntry>(entity =>
        {
            entity.HasIndex(x => x.PlannedStart);
            entity.HasIndex(x => x.EquipmentId);
            entity.HasIndex(x => x.EmployeeId);
            entity.HasIndex(x => x.DemandItemId);
        });

        modelBuilder.Entity<AppUser>(entity =>
        {
            entity.Property(x => x.UserName).HasMaxLength(64);
            entity.Property(x => x.NormalizedUserName).HasMaxLength(64);
            entity.Property(x => x.DisplayName).HasMaxLength(256);
            entity.Property(x => x.PasswordHash).HasMaxLength(256);
            entity.HasIndex(x => x.NormalizedUserName).IsUnique();
        });

        modelBuilder.Entity<AppUserPermission>(entity =>
        {
            entity.Property(x => x.PageKey).HasMaxLength(64);
            entity.HasIndex(x => new { x.AppUserId, x.PageKey }).IsUnique();
        });

        modelBuilder.Entity<AuthLoginAttempt>(entity =>
        {
            entity.Property(x => x.UserName).HasMaxLength(64);
            entity.Property(x => x.NormalizedUserName).HasMaxLength(64);
            entity.Property(x => x.FailureReason).HasMaxLength(256);
            entity.Property(x => x.MachineName).HasMaxLength(128);
            entity.HasIndex(x => new { x.NormalizedUserName, x.AttemptedAt });
        });

        foreach (var entityType in modelBuilder.Model.GetEntityTypes())
        {
            foreach (var property in entityType.GetProperties().Where(p => p.ClrType == typeof(decimal) || p.ClrType == typeof(decimal?)))
            {
                property.SetPrecision(18);
                property.SetScale(4);
            }
        }
    }

    public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default)
    {
        AddChangeHistory();
        return base.SaveChangesAsync(cancellationToken);
    }

    private void AddChangeHistory()
    {
        var audited = new[] { nameof(Part), nameof(CanonicalBlank), nameof(BlankAlias), nameof(PartBlankMap) };
        var entries = ChangeTracker.Entries()
            .Where(e => audited.Contains(e.Entity.GetType().Name) && e.State == EntityState.Modified)
            .ToList();

        foreach (var entry in entries)
        {
            var idProperty = entry.Property(nameof(Entity.Id));
            if (idProperty.CurrentValue is not long id || id == 0)
            {
                continue;
            }

            foreach (var property in entry.Properties.Where(p => p.IsModified))
            {
                ChangeHistory.Add(new ChangeHistory
                {
                    EntityType = entry.Entity.GetType().Name,
                    EntityId = id,
                    FieldName = property.Metadata.Name,
                    OldValue = property.OriginalValue?.ToString(),
                    NewValue = property.CurrentValue?.ToString(),
                    Source = "Application"
                });
            }
        }
    }
}
