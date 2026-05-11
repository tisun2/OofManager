using System.IO;
using System.Text;

namespace OofManager.Wpf.Services;

/// <summary>
/// Append-only diagnostic log for background auto-sync activity. We write
/// to %LOCALAPPDATA%\OofManager\sync.log so that when the user reports
/// "OofManager and Outlook disagree", we can read this file and see exactly
/// what OofManager did, what the server returned before / after each push,
/// and any exception messages from a background tick — failures in the
/// 5-minute auto-sync loop only surface in the in-app status bar, which the
/// user can't see when the window is minimized to the tray.
/// </summary>
public static class SyncLogger
{
    private static readonly object _lock = new();
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OofManager",
        "sync.log");
    // Cap the file at ~1MB. When we cross that, rotate to .1 (overwriting the
    // previous rotation). One generation is enough for "what happened in the
    // last few hours"; we don't need full archival history.
    private const long MaxBytes = 1024 * 1024;

    public static void Write(string line)
    {
        try
        {
            lock (_lock)
            {
                var dir = Path.GetDirectoryName(LogPath)!;
                Directory.CreateDirectory(dir);
                RotateIfNeeded();
                var stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                File.AppendAllText(LogPath, $"[{stamp}] {line}{Environment.NewLine}", Encoding.UTF8);
            }
        }
        catch
        {
            // Diagnostics must never break the app. Disk full, permission
            // denied, antivirus lock — all silently ignored.
        }
    }

    private static void RotateIfNeeded()
    {
        try
        {
            var info = new FileInfo(LogPath);
            if (!info.Exists || info.Length < MaxBytes) return;
            var rotated = LogPath + ".1";
            if (File.Exists(rotated)) File.Delete(rotated);
            File.Move(LogPath, rotated);
        }
        catch
        {
            // Best-effort rotation; if it fails the next Append will just
            // grow the file slightly past the cap until it succeeds.
        }
    }
}
