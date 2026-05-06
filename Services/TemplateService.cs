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
        await _db.CreateTableAsync<OofTemplate>();
        return _db;
    }

    public async Task<List<OofTemplate>> GetAllTemplatesAsync()
    {
        var db = await GetDbAsync();
        return await db.Table<OofTemplate>().OrderByDescending(t => t.UpdatedAt).ToListAsync();
    }

    public async Task<OofTemplate?> GetTemplateAsync(int id)
    {
        var db = await GetDbAsync();
        return await db.FindAsync<OofTemplate>(id);
    }

    public async Task SaveTemplateAsync(OofTemplate template)
    {
        var db = await GetDbAsync();
        template.UpdatedAt = DateTime.UtcNow;

        if (template.Id == 0)
        {
            template.CreatedAt = DateTime.UtcNow;
            await db.InsertAsync(template);
        }
        else
        {
            await db.UpdateAsync(template);
        }
    }

    public async Task DeleteTemplateAsync(int id)
    {
        var db = await GetDbAsync();
        await db.DeleteAsync<OofTemplate>(id);
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
