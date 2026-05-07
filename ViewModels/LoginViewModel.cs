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

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private string _statusMessage = "Sign in with your Microsoft 365 account";

    [ObservableProperty]
    private bool _isLoggedIn;

    [ObservableProperty]
    private string _userDisplayName = string.Empty;

    private bool _autoLoginAttempted;

    public LoginViewModel(IExchangeService exchangeService, INavigationService navigation, IPreferencesService prefs)
    {
        _exchangeService = exchangeService;
        _navigation = navigation;
        _prefs = prefs;
        // App.OnStartup already kicked off PrewarmAsync; calling it here would be
        // a no-op anyway because PrewarmAsync caches the in-flight task.
    }

    /// <summary>
    /// Called when the LoginPage becomes visible. If we have a remembered UPN
    /// from a previous successful sign-in, ask MSAL/WAM for a silent token via
    /// Connect-ExchangeOnline -UserPrincipalName. When the token cache is warm
    /// (typical on Windows for ~90 days after the last interactive sign-in),
    /// this skips the LoginPage entirely. Falls back silently to the manual
    /// Sign In button if anything goes wrong.
    /// </summary>
    public async Task TryAutoLoginAsync()
    {
        if (_autoLoginAttempted) return;
        _autoLoginAttempted = true;

        var lastUpn = _prefs.GetString(LastUpnPrefKey);
        if (string.IsNullOrWhiteSpace(lastUpn)) return;
        if (IsBusy || _exchangeService.IsConnected) return;

        IsBusy = true;
        StatusMessage = $"Signing you in as {lastUpn}…";

        try
        {
            await _exchangeService.ConnectAsync(upnHint: lastUpn);
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
            // Pass the last-known UPN so WAM defaults to the same account; if
            // the user wants to switch accounts they can still pick "Use a
            // different account" in the WAM dialog.
            var lastUpn = _prefs.GetString(LastUpnPrefKey);
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
