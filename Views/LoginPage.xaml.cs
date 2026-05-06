using System.Windows.Controls;
using OofManager.Wpf.ViewModels;

namespace OofManager.Wpf.Views;

public partial class LoginPage : Page
{
    public LoginPage(LoginViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }
}
