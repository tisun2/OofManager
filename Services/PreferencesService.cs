using System.IO;
using System.Text.Json;

namespace OofManager.Wpf.Services;

public class PreferencesService : IPreferencesService
{
    private static readonly string LocalFallbackPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OofManager",
        "preferences.json");

    private static readonly string FilePath = ResolvePreferencesFilePath();

    /// <summary>
    /// Resolves the synced storage directory, preferring OneDrive for Business
    /// (<c>%OneDriveCommercial%\OofManager</c>) or personal OneDrive
    /// (<c>%OneDrive%\OofManager</c>) when present on this machine, and falling
    /// back to <c>%LOCALAPPDATA%\OofManager</c> otherwise.
    /// </summary>
    public static string ResolveSyncedDataDirectory(string? oneDriveRootOverride = null, string? localAppDataOverride = null)
    {
        var candidates = oneDriveRootOverride != null
            ? new[] { oneDriveRootOverride }
            : new[]
            {
                Environment.GetEnvironmentVariable("OneDriveCommercial"),
                Environment.GetEnvironmentVariable("OneDrive"),
                Environment.GetEnvironmentVariable("OneDriveConsumer"),
            };

        foreach (var candidate in candidates)
        {
            if (string.IsNullOrWhiteSpace(candidate)) continue;
            var trimmed = candidate!.Trim();
            try
            {
                if (Directory.Exists(trimmed))
                    return Path.Combine(trimmed, "OofManager");
            }
            catch { }
        }

        var localRoot = !string.IsNullOrWhiteSpace(localAppDataOverride)
            ? localAppDataOverride!
            : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(localRoot, "OofManager");
    }

    public static string ResolvePreferencesFilePath(string? oneDriveRootOverride = null, string? localAppDataOverride = null)
    {
        var syncedDir = ResolveSyncedDataDirectory(oneDriveRootOverride, localAppDataOverride);
        var targetPath = Path.Combine(syncedDir, "preferences.json");
        var localRoot = !string.IsNullOrWhiteSpace(localAppDataOverride)
            ? localAppDataOverride!
            : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var legacyLocalPath = Path.Combine(localRoot, "OofManager", "preferences.json");

        // One-time seamless migration: if OneDrive is active and doesn't have
        // preferences.json yet, seed it from %LOCALAPPDATA%\OofManager\preferences.json.
        if (!string.Equals(targetPath, legacyLocalPath, StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                if (!File.Exists(targetPath) && File.Exists(legacyLocalPath))
                {
                    Directory.CreateDirectory(syncedDir);
                    File.Copy(legacyLocalPath, targetPath, overwrite: false);
                }
            }
            catch { }
        }

        return targetPath;
    }

    private Dictionary<string, JsonElement> _data = new();
    private bool _loaded;
    private readonly object _lock = new();
    private int _batchDepth;
    private bool _dirtyDuringBatch;

    private void EnsureLoaded()
    {
        if (_loaded) return;
        lock (_lock)
        {
            if (_loaded) return;
            try
            {
                var pathToRead = File.Exists(FilePath)
                    ? FilePath
                    : (File.Exists(LocalFallbackPath) ? LocalFallbackPath : FilePath);
                if (File.Exists(pathToRead))
                {
                    var json = File.ReadAllText(pathToRead);
                    _data = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json) ?? new();
                }
            }
            catch { _data = new(); }
            _loaded = true;
        }
    }

    private void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            // Snapshot under the lock so we serialize a consistent view even if a
            // concurrent Set() races in.
            string json;
            lock (_lock)
            {
                json = JsonSerializer.Serialize(_data);
            }
            // Atomic write: write to a temp file, then replace the target. This keeps the
            // existing file intact if the process is killed / crashes mid-write — otherwise
            // a truncated JSON file silently resets all the user's preferences on next load.
            var tmp = FilePath + ".tmp";
            File.WriteAllText(tmp, json);
            if (File.Exists(FilePath))
            {
                // File.Replace is atomic on NTFS and preserves attributes/ACLs.
                File.Replace(tmp, FilePath, destinationBackupFileName: null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tmp, FilePath);
            }
        }
        catch { }
    }

    public bool GetBool(string key, bool defaultValue)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (_data.TryGetValue(key, out var v))
            {
                if (v.ValueKind == JsonValueKind.True) return true;
                if (v.ValueKind == JsonValueKind.False) return false;
            }
        }
        return defaultValue;
    }

    public int GetInt(string key, int defaultValue)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (_data.TryGetValue(key, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var i))
                return i;
        }
        return defaultValue;
    }

    public string? GetString(string key, string? defaultValue = null)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (_data.TryGetValue(key, out var v) && v.ValueKind == JsonValueKind.String)
                return v.GetString();
        }
        return defaultValue;
    }

    public void Set(string key, bool value)
    {
        EnsureLoaded();
        lock (_lock)
        {
            // Skip the write if the value is unchanged.
            if (_data.TryGetValue(key, out var existing))
            {
                if (existing.ValueKind == (value ? JsonValueKind.True : JsonValueKind.False)) return;
            }
            _data[key] = JsonSerializer.SerializeToElement(value);
        }
        TrySave();
    }

    public void Set(string key, int value)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (_data.TryGetValue(key, out var existing) && existing.ValueKind == JsonValueKind.Number
                && existing.TryGetInt32(out var current) && current == value)
                return;
            _data[key] = JsonSerializer.SerializeToElement(value);
        }
        TrySave();
    }

    public void Set(string key, string? value)
    {
        EnsureLoaded();
        lock (_lock)
        {
            // Treat null/empty as "remove" so callers don't accumulate dead keys
            // (e.g. after a sign-out clearing the last UPN).
            if (string.IsNullOrEmpty(value))
            {
                if (!_data.Remove(key)) return;
            }
            else
            {
                if (_data.TryGetValue(key, out var existing) && existing.ValueKind == JsonValueKind.String
                    && string.Equals(existing.GetString(), value, StringComparison.Ordinal))
                    return;
                _data[key] = JsonSerializer.SerializeToElement(value);
            }
        }
        TrySave();
    }

    private void TrySave()
    {
        if (_batchDepth > 0)
        {
            _dirtyDuringBatch = true;
            return;
        }
        Save();
    }

    public IDisposable BeginBatch() => new BatchScope(this);

    private sealed class BatchScope : IDisposable
    {
        private readonly PreferencesService _owner;
        private bool _disposed;

        public BatchScope(PreferencesService owner)
        {
            _owner = owner;
            _owner._batchDepth++;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _owner._batchDepth--;
            if (_owner._batchDepth == 0 && _owner._dirtyDuringBatch)
            {
                _owner._dirtyDuringBatch = false;
                _owner.Save();
            }
        }
    }
}
