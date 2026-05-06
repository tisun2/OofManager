using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OofManager.Wpf.Services;

namespace OofManager.Wpf.ViewModels;

public partial class LoginViewModel : ObservableObject
{
    private readonly IExchangeService _exchangeService;
    private readonly INavigationService _navigation;

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private string _statusMessage = "Sign in with your Microsoft 365 account";

    [ObservableProperty]
    private bool _isLoggedIn;

    [ObservableProperty]
    private string _userDisplayName = string.Empty;

    public LoginViewModel(IExchangeService exchangeService, INavigationService navigation)
    {
        _exchangeService = exchangeService;
        _navigation = navigation;
        // App.OnStartup already kicked off PrewarmAsync; calling it here would be
        // a no-op anyway because PrewarmAsync caches the in-flight task.
    }

    [RelayCommand]
    private async Task LoginAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        StatusMessage = "Signing in to Microsoft 365... (a sign-in dialog will appear)";

        try
        {
            await _exchangeService.ConnectAsync();
            IsLoggedIn = true;
            UserDisplayName = await _exchangeService.GetCurrentUserAsync();
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
