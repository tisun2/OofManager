using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace OofManager.Wpf.Services;

public interface IStartupService
{
    /// <summary>True if OofManager is currently registered to launch at user logon.</summary>
    bool IsEnabled { get; }

    /// <summary>
    /// True if the autostart entry exists AND points at the currently-running
    /// executable. False if the entry is missing or points somewhere else
    /// (e.g. a stale dev build path left over from a previous install).
    /// </summary>
    bool IsRegisteredForCurrentExe { get; }

    /// <summary>
    /// Register (or unregister) OofManager for user-scoped autostart by writing
    /// the HKCU Run key. When enabled, the launch command includes the
    /// <c>--minimized</c> switch so the app comes up hidden in the tray instead
    /// of stealing focus.
    /// </summary>
    void SetEnabled(bool enabled);

    /// <summary>
    /// If autostart is currently enabled but the registered command points at
    /// a different executable (e.g. an old dev build path that survived an
    /// install), rewrite the entry to point at the current exe. This keeps the
    /// user's "start with Windows" choice working across upgrades / reinstalls
    /// without ever prompting them again.
    /// </summary>
    void EnsureRegistrationIsFresh();

    /// <summary>True if the user has already been shown the autostart prompt
    /// for the currently-installed version. Returns false after an upgrade so
    /// the user gets one chance to reconsider on each new version (the older
    /// per-profile flag silently buried this question across upgrades).</summary>
    bool HasBeenPromptedBefore();

    /// <summary>Records that the autostart prompt has been shown for the
    /// currently-installed version. Should only be called once the dialog has
    /// actually been displayed to the user.</summary>
    void MarkPromptShown();
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
    // Legacy boolean flag ("prompt has ever been shown") — honoured for
    // back-compat so users who said "yes" on a previous version don't get
    // re-asked. The per-version key below is what new code reads/writes.
    private const string PromptPrefKey = "Startup.PromptShown";
    // Stores the app version (e.g. "1.1.0") for which we last showed the
    // autostart prompt. After an upgrade this no longer matches the running
    // version, so HasBeenPromptedBefore() returns false and the user gets one
    // fresh chance to opt in.
    private const string PromptVersionPrefKey = "Startup.PromptShownVersion";
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

    public bool IsRegisteredForCurrentExe
    {
        get
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
                var value = key?.GetValue(RunValueName) as string;
                if (string.IsNullOrWhiteSpace(value)) return false;

                var current = GetExecutablePath();
                if (string.IsNullOrEmpty(current)) return false;

                // Stored command is `"<exe>" --minimized`; pull the exe path out
                // of the leading quoted segment (or, defensively, the first whitespace-
                // delimited token if quoting was ever stripped).
                var registered = ExtractExePath(value!);
                return !string.IsNullOrEmpty(registered)
                    && string.Equals(
                        Path.GetFullPath(registered!),
                        Path.GetFullPath(current!),
                        StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
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

    public void EnsureRegistrationIsFresh()
    {
        // Self-heal a stale Run entry. Without this, an autostart enabled from
        // a previous install (e.g. a dev build at bin\Debug\…) survives into a
        // fresh install and Windows tries to launch the old, possibly-missing
        // exe at logon — manifesting as "I enabled it but it doesn't start".
        try
        {
            if (!IsEnabled) return;            // user opted out → nothing to do
            if (IsRegisteredForCurrentExe) return; // already pointing at us
            SetEnabled(true);                  // rewrite with current exe path
        }
        catch
        {
            // Non-fatal: worst case the user re-toggles the checkbox manually.
        }
    }

    public bool HasBeenPromptedBefore()
    {
        // If autostart is currently *enabled*, the user already opted in at
        // some point — don't re-ask just because they upgraded.
        if (IsEnabled) return true;

        var current = GetCurrentAppVersion();
        var lastPromptedVersion = _prefs.GetString(PromptVersionPrefKey, null);

        // Prompt once per installed version. If we already showed the prompt
        // for *this* version, stay quiet. After an upgrade the value no longer
        // matches and the user gets one fresh chance to opt in.
        return !string.IsNullOrEmpty(lastPromptedVersion)
            && !string.IsNullOrEmpty(current)
            && string.Equals(lastPromptedVersion, current, StringComparison.OrdinalIgnoreCase);
        // NOTE: deliberately ignore the legacy PromptPrefKey boolean. Its old
        // semantics ("prompt has ever been shown, ever") permanently buried
        // the question for users who declined once on an older version, even
        // across reinstalls and upgrades. New behaviour is per-version, and
        // the one-time re-prompt those users get on first launch after this
        // change is the intended fix.
    }

    public void MarkPromptShown()
    {
        _prefs.Set(PromptPrefKey, true);
        var current = GetCurrentAppVersion();
        if (!string.IsNullOrEmpty(current)) _prefs.Set(PromptVersionPrefKey, current);
    }

    private static string? GetCurrentAppVersion()
    {
        try
        {
            var path = Process.GetCurrentProcess().MainModule?.FileName;
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return null;
            // FileVersionInfo.FileVersion matches the AssemblyFileVersion baked
            // in by MSBuild and bumped by the installer pipeline, so it ticks
            // forward on every release.
            return FileVersionInfo.GetVersionInfo(path).FileVersion;
        }
        catch
        {
            return null;
        }
    }

    private static string? ExtractExePath(string command)
    {
        if (string.IsNullOrWhiteSpace(command)) return null;
        var trimmed = command.TrimStart();
        if (trimmed.StartsWith("\""))
        {
            var end = trimmed.IndexOf('"', 1);
            if (end > 1) return trimmed.Substring(1, end - 1);
        }
        // Fallback for legacy unquoted entries — split on first whitespace.
        var space = trimmed.IndexOf(' ');
        return space < 0 ? trimmed : trimmed.Substring(0, space);
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
