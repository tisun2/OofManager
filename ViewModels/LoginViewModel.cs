using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OofManager.Wpf.Services;

namespace OofManager.Wpf.ViewModels;

public partial class LoginViewModel : ObservableObject
{
    // Stored across launches so we can ask MSAL/WAM for a silent token-cache
    // refresh on the next start instead of forcing the user to click Sign In.
    private const string LastUpnPrefKey = "Auth.LastSignedInUpn";

    private readonly IExchangeService _exchangeService;
    private readonly INavigationService _navigation;
    private readonly IPreferencesService _prefs;
    private readonly IWindowsAccountService _windowsAccount;

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private string _statusMessage = "Sign in with your Microsoft 365 account";

    [ObservableProperty]
    private bool _isLoggedIn;

    [ObservableProperty]
    private string _userDisplayName = string.Empty;

    private bool _autoLoginAttempted;

    public LoginViewModel(
        IExchangeService exchangeService,
        INavigationService navigation,
        IPreferencesService prefs,
        IWindowsAccountService windowsAccount)
    {
        _exchangeService = exchangeService;
        _navigation = navigation;
        _prefs = prefs;
        _windowsAccount = windowsAccount;
        // App.OnStartup already kicked off PrewarmAsync; calling it here would be
        // a no-op anyway because PrewarmAsync caches the in-flight task.
    }

    /// <summary>
    /// Called when the LoginPage becomes visible. Tries to sign the user in
    /// without any UI:
    ///   1. If a previous successful sign-in cached a UPN, use it.
    ///   2. Otherwise fall back to the Windows account UPN — on an Entra-joined
    ///      device this lets first-ever launch complete via WAM SSO with no
    ///      account picker and no prompt.
    /// In both cases the cached MSAL refresh token (or the device-bound PRT on
    /// Entra-joined machines) does the heavy lifting; if there's no usable
    /// token, WAM falls back to its interactive dialog and we bail out so the
    /// user sees the manual Sign In button instead.
    /// </summary>
    public async Task TryAutoLoginAsync()
    {
        if (_autoLoginAttempted) return;
        _autoLoginAttempted = true;

        // Prefer the UPN we explicitly remembered from a prior sign-in; this
        // honors users who picked a non-default account in WAM. Only fall back
        // to the Windows UPN when we have nothing better — that's the case on
        // a brand-new install on a corporate Entra-joined device.
        var upnHint = _prefs.GetString(LastUpnPrefKey);
        var fromWindows = false;
        if (string.IsNullOrWhiteSpace(upnHint))
        {
            upnHint = _windowsAccount.TryGetCurrentUserUpn();
            fromWindows = !string.IsNullOrWhiteSpace(upnHint);
        }

        if (string.IsNullOrWhiteSpace(upnHint)) return;
        if (IsBusy || _exchangeService.IsConnected) return;

        IsBusy = true;
        StatusMessage = fromWindows
            ? $"Signing you in as {upnHint} (Windows account)…"
            : $"Signing you in as {upnHint}…";

        try
        {
            await _exchangeService.ConnectAsync(upnHint: upnHint);
            IsLoggedIn = true;
            UserDisplayName = await _exchangeService.GetCurrentUserAsync();
            // Refresh the cached UPN in case the canonical form differs from
            // what the user originally typed (e.g. casing, alias vs. UPN).
            _prefs.Set(LastUpnPrefKey, UserDisplayName);
            StatusMessage = $"Welcome back, {UserDisplayName}!";
            _navigation.NavigateToMain();
        }
        catch
        {
            // Silent auto-login failed (token expired, password change, CA
            // policy, network blip, etc.). Drop back to the Sign In button —
            // the user can click it and go through interactive auth as before.
            StatusMessage = "Sign in with your Microsoft 365 account";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task LoginAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        StatusMessage = "Signing in to Microsoft 365... (a sign-in dialog will appear)";

        try
        {
            // Pass the last-known UPN (or Windows UPN as fallback) so WAM
            // defaults to the same account; the user can still pick "Use a
            // different account" inside the WAM dialog if they want to switch.
            var lastUpn = _prefs.GetString(LastUpnPrefKey)
                          ?? _windowsAccount.TryGetCurrentUserUpn();
            await _exchangeService.ConnectAsync(upnHint: lastUpn);
            IsLoggedIn = true;
            UserDisplayName = await _exchangeService.GetCurrentUserAsync();
            _prefs.Set(LastUpnPrefKey, UserDisplayName);
            StatusMessage = $"Welcome, {UserDisplayName}!";
            _navigation.NavigateToMain();
        }
        catch (Exception ex)
        {
            StatusMessage = $"Sign-in failed: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }
}
