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

    public void NavigateToLogin(bool forceAccountPicker = false)
    {
        var page = _services.GetService(typeof(LoginPage)) as LoginPage
                   ?? throw new InvalidOperationException("LoginPage not registered");
        _frame.Navigate(page);

        if (forceAccountPicker)
        {
            // Defer to Background priority so the Frame finishes its navigation
            // and the LoginPage is actually on screen before the Switch Account
            // prompt dialog appears. Otherwise the dialog would briefly overlay
            // MainPage while the Frame is still tearing it down.
            var vm = _services.GetService(typeof(ViewModels.LoginViewModel)) as ViewModels.LoginViewModel;
            if (vm != null)
            {
                System.Windows.Application.Current.Dispatcher.BeginInvoke(
                    new Action(async () => await vm.SwitchAccountAsync()),
                    System.Windows.Threading.DispatcherPriority.Background);
            }
        }
    }
}
