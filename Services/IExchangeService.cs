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

    Task ConnectAsync();
    Task DisconnectAsync();
    Task<OofSettings> GetOofSettingsAsync();
    Task SetOofSettingsAsync(OofSettings settings);
    Task<string> GetCurrentUserAsync();
    bool IsConnected { get; }
}
