using System.IO;
using OofManager.Wpf.Models;
using SQLite;

namespace OofManager.Wpf.Services;

public class TemplateService : ITemplateService, IAsyncDisposable
{
    static TemplateService()
    {
        // Initialize the SQLite native provider (required on .NET Framework).
        SQLitePCL.Batteries_V2.Init();
    }

    private SQLiteAsyncConnection? _db;

    private async Task<SQLiteAsyncConnection> GetDbAsync()
    {
        if (_db != null) return _db;

        var appDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OofManager");
        Directory.CreateDirectory(appDataDir);
        var dbPath = Path.Combine(appDataDir, "oofmanager.db3");
        _db = new SQLiteAsyncConnection(dbPath);
        await _db.CreateTableAsync<OofTemplate>().ConfigureAwait(false);
        return _db;
    }

    public async Task<List<OofTemplate>> GetAllTemplatesAsync()
    {
        var db = await GetDbAsync().ConfigureAwait(false);
        return await db.Table<OofTemplate>().OrderByDescending(t => t.UpdatedAt).ToListAsync().ConfigureAwait(false);
    }

    public async Task<OofTemplate?> GetTemplateAsync(int id)
    {
        var db = await GetDbAsync().ConfigureAwait(false);
        return await db.FindAsync<OofTemplate>(id).ConfigureAwait(false);
    }

    public async Task SaveTemplateAsync(OofTemplate template)
    {
        var db = await GetDbAsync().ConfigureAwait(false);
        template.UpdatedAt = DateTime.UtcNow;

        if (template.Id == 0)
        {
            template.CreatedAt = DateTime.UtcNow;
            await db.InsertAsync(template).ConfigureAwait(false);
        }
        else
        {
            await db.UpdateAsync(template).ConfigureAwait(false);
        }
    }

    public async Task DeleteTemplateAsync(int id)
    {
        var db = await GetDbAsync().ConfigureAwait(false);
        await db.DeleteAsync<OofTemplate>(id).ConfigureAwait(false);
    }

    /// <summary>
    /// Closes the underlying SQLite connection so any pending WAL is flushed cleanly.
    /// Called by the DI container on application exit.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_db != null)
        {
            try { await _db.CloseAsync(); } catch { }
            _db = null;
        }
    }
}
