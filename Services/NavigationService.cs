using System.Windows.Controls;
using OofManager.Wpf.Views;

namespace OofManager.Wpf.Services;

public class NavigationService : INavigationService
{
    private readonly Frame _frame;
    private readonly IServiceProvider _services;

    public NavigationService(Frame frame, IServiceProvider services)
    {
        _frame = frame;
        _services = services;
    }

    public void NavigateToMain()
    {
        var page = _services.GetService(typeof(MainPage)) as MainPage
                   ?? throw new InvalidOperationException("MainPage not registered");
        _frame.Navigate(page);
    }

    public void NavigateToLogin()
    {
        var page = _services.GetService(typeof(LoginPage)) as LoginPage
                   ?? throw new InvalidOperationException("LoginPage not registered");
        _frame.Navigate(page);
    }
}
