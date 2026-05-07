using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace OofManager.Wpf.Services;

public interface IStartupService
{
    /// <summary>True if OofManager is currently registered to launch at user logon.</summary>
    bool IsEnabled { get; }

    /// <summary>
    /// Register (or unregister) OofManager for user-scoped autostart by writing
    /// the HKCU Run key. When enabled, the launch command includes the
    /// <c>--minimized</c> switch so the app comes up hidden in the tray instead
    /// of stealing focus.
    /// </summary>
    void SetEnabled(bool enabled);

    /// <summary>Returns true the very first time the app should ask the user
    /// about autostart, false on every subsequent run. Records the prompt as
    /// shown, so no caller needs to track the state itself.</summary>
    bool ShouldPromptUser();
}

/// <summary>
/// User-scoped autostart via HKCU\Software\Microsoft\Windows\CurrentVersion\Run.
/// Per-user registration needs no admin elevation and roams with the profile.
/// We write the *current* process's main-module path so the same machine that
/// installed the app is the one that launches it at login (no broken paths
/// after upgrades that move the binary; the next manual launch fixes the entry).
/// </summary>
public sealed class StartupService : IStartupService
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string RunValueName = "OofManager";
    private const string PromptPrefKey = "Startup.PromptShown";
    public const string MinimizedArg = "--minimized";

    private readonly IPreferencesService _prefs;

    public StartupService(IPreferencesService prefs)
    {
        _prefs = prefs;
    }

    public bool IsEnabled
    {
        get
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
                var value = key?.GetValue(RunValueName) as string;
                return !string.IsNullOrWhiteSpace(value);
            }
            catch
            {
                // Group policy / locked-down profile: treat as not enabled and
                // don't surface the registry exception to the user.
                return false;
            }
        }
    }

    public void SetEnabled(bool enabled)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true);
            if (key == null) return;
            if (enabled)
            {
                var exePath = GetExecutablePath();
                if (string.IsNullOrEmpty(exePath)) return;
                // Always quote the path; spaces in "Program Files" or in the
                // user profile would otherwise truncate the command.
                key.SetValue(RunValueName, $"\"{exePath}\" {MinimizedArg}", RegistryValueKind.String);
            }
            else
            {
                key.DeleteValue(RunValueName, throwOnMissingValue: false);
            }
        }
        catch
        {
            // Failing to write the Run key is non-fatal — UI will reflect the
            // real state on next read.
        }
    }

    public bool ShouldPromptUser()
    {
        // One-shot: only the first session ever asks. We deliberately gate on
        // a pref instead of "is the Run key empty" so that a user who said No
        // isn't re-asked on every launch.
        if (_prefs.GetBool(PromptPrefKey, false)) return false;
        _prefs.Set(PromptPrefKey, true);
        return true;
    }

    private static string? GetExecutablePath()
    {
        // MainModule.FileName works under both regular launches and dotnet-host
        // launches because OofManager.Wpf is a real .NET Framework exe (net48).
        try
        {
            var path = Process.GetCurrentProcess().MainModule?.FileName;
            return string.IsNullOrEmpty(path) || !File.Exists(path) ? null : path;
        }
        catch
        {
            return null;
        }
    }
}
