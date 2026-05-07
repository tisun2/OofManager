using System.Windows.Controls;
using OofManager.Wpf.ViewModels;

namespace OofManager.Wpf.Views;

public partial class LoginPage : Page
{
    private readonly LoginViewModel _viewModel;

    public LoginPage(LoginViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;
        // Kick off the silent-token auto-login as soon as the page is on screen.
        // Loaded fires after Frame navigation; the VM guards against running it
        // more than once per session.
        Loaded += async (_, _) => await _viewModel.TryAutoLoginAsync();
    }
}
