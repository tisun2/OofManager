using OofManager.Wpf.Models;

namespace OofManager.Wpf.Services;

public interface IExchangeService
{
    /// <summary>
    /// Pre-warms the PowerShell host and pre-loads ExchangeOnlineManagement so that
    /// a subsequent ConnectAsync only has to perform the user authentication step.
    /// Safe to call multiple times; subsequent calls are no-ops.
    /// </summary>
    Task PrewarmAsync();

    /// <summary>
    /// Connects to Exchange Online. When <paramref name="upnHint"/> is provided,
    /// the UPN is passed to Connect-ExchangeOnline so MSAL/WAM can attempt a
    /// silent token-cache hit (no UI) if a fresh refresh token is already
    /// cached for that identity. When null, WAM shows its account picker.
    /// </summary>
    Task ConnectAsync(string? upnHint = null);

    /// <summary>
    /// Best-effort silent reconnect: kicks off (or reuses) a background
    /// <see cref="ConnectAsync"/> with the given UPN hint and caches the resulting
    /// task so concurrent callers share it. Exceptions are swallowed inside the
    /// task; callers must inspect <see cref="IsConnected"/> after awaiting to
    /// decide whether the silent attempt actually succeeded. Designed to be
    /// fired during app startup so the LoginPage's auto-sign-in path observes
    /// IsConnected==true and skips straight to MainPage in the common case.
    /// </summary>
    Task TryAutoConnectAsync(string upnHint);
    Task DisconnectAsync();
    Task<OofSettings> GetOofSettingsAsync();
    Task SetOofSettingsAsync(OofSettings settings);
    Task<string> GetCurrentUserAsync();
    Task<string> GetCurrentMailboxIdentityAsync();
    bool IsConnected { get; }
}
