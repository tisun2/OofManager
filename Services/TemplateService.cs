using System.IO;
using System.Text.Json;
using OofManager.Wpf.Models;
using SQLite;

namespace OofManager.Wpf.Services;

/// <summary>
/// Stores OOF reply templates in a OneDrive-synced JSON file (<c>templates.json</c>)
/// so templates roam seamlessly across machines without SQLite WAL/file-lock
/// conflicts on OneDrive. Automatically imports any existing templates from the
/// legacy <c>%LOCALAPPDATA%\OofManager\oofmanager.db3</c> SQLite database on first
/// run when <c>templates.json</c> does not exist yet.
/// </summary>
public class TemplateService : ITemplateService, IAsyncDisposable
{
    private readonly string _jsonFilePath;
    private readonly string _legacyDbPath;
    private readonly object _lock = new();
    private List<OofTemplate> _templates = new();
    private bool _loaded;
    private DateTime _lastLoadedWriteUtc = DateTime.MinValue;

    public TemplateService()
        : this(
            Path.Combine(PreferencesService.ResolveSyncedDataDirectory(), "templates.json"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OofManager", "oofmanager.db3"))
    {
    }

    public TemplateService(string jsonFilePath, string? legacyDbPath = null)
    {
        _jsonFilePath = jsonFilePath;
        _legacyDbPath = legacyDbPath ?? string.Empty;
    }

    private void EnsureLoaded()
    {
        lock (_lock)
        {
            if (File.Exists(_jsonFilePath))
            {
                try
                {
                    var writeUtc = File.GetLastWriteTimeUtc(_jsonFilePath);
                    if (_loaded && writeUtc == _lastLoadedWriteUtc) return;
                    var json = File.ReadAllText(_jsonFilePath);
                    _templates = JsonSerializer.Deserialize<List<OofTemplate>>(json) ?? new();
                    _lastLoadedWriteUtc = writeUtc;
                    _loaded = true;
                    return;
                }
                catch
                {
                    if (_loaded) return;
                    _templates = new();
                }
            }

            if (_loaded) return;
            _loaded = true;

            // One-time migration from legacy SQLite database (oofmanager.db3) if present.
            if (!string.IsNullOrWhiteSpace(_legacyDbPath) && File.Exists(_legacyDbPath))
            {
                try
                {
                    SQLitePCL.Batteries_V2.Init();
                    using var conn = new SQLiteConnection(_legacyDbPath, SQLiteOpenFlags.ReadOnly);
                    _templates = conn.Table<OofTemplate>()
                        .OrderByDescending(t => t.UpdatedAt)
                        .ToList();
                    SaveLocked();
                }
                catch
                {
                    _templates = new();
                }
            }
        }
    }

    private void SaveLocked()
    {
        try
        {
            var dir = Path.GetDirectoryName(_jsonFilePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir!);

            var json = JsonSerializer.Serialize(_templates, new JsonSerializerOptions { WriteIndented = true });
            var tmp = _jsonFilePath + ".tmp";
            File.WriteAllText(tmp, json);
            if (File.Exists(_jsonFilePath))
            {
                File.Replace(tmp, _jsonFilePath, destinationBackupFileName: null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tmp, _jsonFilePath);
            }
            _lastLoadedWriteUtc = File.GetLastWriteTimeUtc(_jsonFilePath);
        }
        catch { }
    }

    public Task<List<OofTemplate>> GetAllTemplatesAsync()
    {
        EnsureLoaded();
        lock (_lock)
        {
            var ordered = _templates
                .OrderByDescending(t => t.UpdatedAt)
                .Select(CloneTemplate)
                .ToList();
            return Task.FromResult(ordered);
        }
    }

    public Task<OofTemplate?> GetTemplateAsync(int id)
    {
        EnsureLoaded();
        lock (_lock)
        {
            var match = _templates.FirstOrDefault(t => t.Id == id);
            return Task.FromResult(match == null ? null : CloneTemplate(match));
        }
    }

    public Task SaveTemplateAsync(OofTemplate template)
    {
        EnsureLoaded();
        lock (_lock)
        {
            var now = DateTime.UtcNow;
            template.UpdatedAt = now;

            if (template.Id == 0)
            {
                template.Id = _templates.Count == 0 ? 1 : _templates.Max(t => t.Id) + 1;
                template.CreatedAt = now;
                _templates.Add(CloneTemplate(template));
            }
            else
            {
                var idx = _templates.FindIndex(t => t.Id == template.Id);
                if (idx >= 0)
                    _templates[idx] = CloneTemplate(template);
                else
                    _templates.Add(CloneTemplate(template));
            }

            SaveLocked();
        }
        return Task.CompletedTask;
    }

    public Task DeleteTemplateAsync(int id)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (_templates.RemoveAll(t => t.Id == id) > 0)
                SaveLocked();
        }
        return Task.CompletedTask;
    }

    private static OofTemplate CloneTemplate(OofTemplate src) => new()
    {
        Id = src.Id,
        Name = src.Name,
        InternalReply = src.InternalReply,
        ExternalReply = src.ExternalReply,
        ExternalAudienceAll = src.ExternalAudienceAll,
        CreatedAt = src.CreatedAt,
        UpdatedAt = src.UpdatedAt,
    };

    public ValueTask DisposeAsync() => default;
}
