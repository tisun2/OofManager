using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OofManager.Wpf.Services;

namespace OofManager.Wpf.ViewModels;

public partial class LoginViewModel : ObservableObject, IDisposable
{
    // Stored across launches so we can ask MSAL/WAM for a silent token-cache
    // refresh on the next start instead of forcing the user to click Sign In.
    private const string LastUpnPrefKey = "Auth.LastSignedInUpn";
    private static readonly TimeSpan AutoLoginConnectTimeout = TimeSpan.FromSeconds(45);
    private static readonly TimeSpan ManualLoginConnectTimeout = TimeSpan.FromSeconds(90);
    private static readonly TimeSpan SlowSignInStatusDelay = TimeSpan.FromSeconds(20);

    private readonly IExchangeService _exchangeService;
    private readonly INavigationService _navigation;
    private readonly IPreferencesService _prefs;
    private readonly IWindowsAccountService _windowsAccount;
    private readonly IDialogService _dialog;

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
        IWindowsAccountService windowsAccount,
        IDialogService dialog)
    {
        _exchangeService = exchangeService;
        _navigation = navigation;
        _prefs = prefs;
        _windowsAccount = windowsAccount;
        _dialog = dialog;
        // App.OnStartup already kicked off PrewarmAsync; calling it here would be
        // a no-op anyway because PrewarmAsync caches the in-flight task.
        // Subscribe so as ExchangeService moves from "Preparing module\u2026" to
        // "Connecting to Microsoft 365\u2026" we surface that to the user instead
        // of leaving them staring at a frozen "Signing you in\u2026" line for ~10s.
        _exchangeService.SignInPhaseChanged += OnSignInPhaseChanged;
    }

    public void Dispose()
    {
        // Unsubscribe so the singleton ExchangeService doesn't keep this VM
        // (and its event handler closure) alive past app shutdown, and to
        // make the lifetime model robust if the VM is ever re-registered as
        // transient/scoped in the future.
        _exchangeService.SignInPhaseChanged -= OnSignInPhaseChanged;
        GC.SuppressFinalize(this);
    }

    private void OnSignInPhaseChanged(string phase)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher == null) return;
        dispatcher.BeginInvoke(new Action(() =>
        {
            // Only overwrite while a sign-in is actually in flight, so success/
            // failure messages set by the command handlers below are not stomped.
            if (IsBusy) StatusMessage = phase;
        }));
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
        if (string.IsNullOrWhiteSpace(upnHint))
        {
            upnHint = _windowsAccount.TryGetCurrentUserUpn();
        }

        if (string.IsNullOrWhiteSpace(upnHint)) return;
        if (IsBusy) return;

        // App.OnStartup may have already finished a silent reconnect by the
        // time we get here. Skip the spinner entirely in that case.
        if (_exchangeService.IsConnected)
        {
            IsLoggedIn = true;
            UserDisplayName = await _exchangeService.GetCurrentUserAsync();
            _prefs.Set(LastUpnPrefKey, UserDisplayName);
            StatusMessage = $"Welcome back, {UserDisplayName}!";
            _navigation.NavigateToMain();
            return;
        }

        // Non-blocking sign-in: we have a UPN to try, so go straight to the main
        // window and let MainViewModel.LoadAsync drive the silent connect in the
        // background. The user sees the fully-populated window immediately instead
        // of watching the login spinner for the entire 8-20s Exchange connect.
        // App.OnStartup already kicked off TryAutoConnectAsync for this UPN, so the
        // session is usually live (or nearly so) by the time LoadAsync awaits it.
        // If the silent connect ultimately fails, LoadAsync navigates back here so
        // the user can sign in interactively.
        _navigation.NavigateToMain();
    }

    [RelayCommand]
    private async Task LoginAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        StatusMessage = _exchangeService.CurrentSignInPhase
            ?? "Signing in to Microsoft 365... (a sign-in dialog will appear)";

        try
        {
            // Pass the last-known UPN (or Windows UPN as fallback) so WAM
            // defaults to the same account; the user can still pick "Use a
            // different account" inside the WAM dialog if they want to switch.
            var lastUpn = _prefs.GetString(LastUpnPrefKey)
                          ?? _windowsAccount.TryGetCurrentUserUpn();
            await AwaitWithSlowStatusAsync(
                _exchangeService.ConnectAsync(upnHint: lastUpn, timeout: ManualLoginConnectTimeout),
                "Still waiting for Microsoft sign-in or Exchange Online. This can happen on slow networks...");
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

    /// <summary>
    /// Triggered by the main page's "Switch account" button. Asks the user
    /// which Microsoft 365 account they want to switch to, then forwards that
    /// UPN to Connect-ExchangeOnline. We can't simply call WAM with no UPN
    /// here because Disconnect-ExchangeOnline does NOT clear the MSAL token
    /// cache — WAM would silently re-issue a token for the previous account
    /// and the user would see no "switch" actually happen. Passing the new UPN
    /// either uses a different cached token (silent) or triggers an
    /// interactive sign-in for a fresh account.
    /// </summary>
    public async Task SwitchAccountAsync()
    {
        // Block TryAutoLoginAsync from racing with us when LoginPage.Loaded fires;
        // otherwise the silent re-login could win and put the user right back
        // into the account they just asked to switch away from.
        _autoLoginAttempted = true;

        if (IsBusy) return;

        var newUpn = await _dialog.PromptAsync(
            title: "Switch account",
            message: "Enter the email address of the Microsoft 365 account you want to switch to:",
            accept: "Sign in",
            cancel: "Cancel",
            placeholder: "name@company.com");

        if (string.IsNullOrWhiteSpace(newUpn))
        {
            // User cancelled — leave them on the login page so they can either
            // close the window or click Sign In manually.
            StatusMessage = "Sign in with your Microsoft 365 account";
            return;
        }

        var trimmedUpn = newUpn!.Trim();
        IsBusy = true;
        StatusMessage = _exchangeService.CurrentSignInPhase
            ?? $"Signing in as {trimmedUpn}\u2026";

        try
        {
            await AwaitWithSlowStatusAsync(
                _exchangeService.ConnectAsync(upnHint: trimmedUpn, timeout: ManualLoginConnectTimeout),
                "Still waiting for Microsoft sign-in or Exchange Online. This can happen on slow networks...");
            IsLoggedIn = true;
            // Use the canonical UPN the server reports back in case the user
            // typed an alias or different casing.
            UserDisplayName = await _exchangeService.GetCurrentUserAsync();
            // Deliberately do NOT persist this UPN as LastSignedInUpn. Switch
            // Account is treated as a one-shot override: next launch goes back
            // to the default flow (cached UPN if any → else Windows account
            // SSO). Without this, switching to a personal account would
            // permanently override the Windows-account fallback for that user
            // even after a restart, which surprised the user.
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

    private async Task AwaitWithSlowStatusAsync(Task task, string slowStatusMessage)
    {
        var slowStatusTask = Task.Delay(SlowSignInStatusDelay);
        var completed = await Task.WhenAny(task, slowStatusTask);
        if (completed == slowStatusTask && IsBusy)
        {
            StatusMessage = slowStatusMessage;
        }

        await task;
    }
}
