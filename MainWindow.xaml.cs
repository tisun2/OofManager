using System.Windows;
using System.Windows.Navigation;
using OofManager.Wpf.Views;

namespace OofManager.Wpf;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        if (v != null) Title = $"OOF Manager v{v.Major}.{v.Minor}.{v.Build}";
    }

    public void Initialize(IServiceProvider services)
    {
        // Always boot into LoginPage. When a UPN is remembered, App.OnStartup
        // has kicked off a silent reconnect in the background and LoginPage's
        // TryAutoLoginAsync awaits it behind a ProgressRing — on success it
        // navigates straight to MainPage. Going directly to MainPage when a
        // UPN is remembered doesn't actually shave any time off the auth wait
        // (we still block on the same Connect-ExchangeOnline call), and it
        // leaves the user stranded with no Sign In button if the silent
        // attempt fails (token expired, password change, network blip).
        var loginPage = (LoginPage)services.GetService(typeof(LoginPage))!;
        RootFrame.Navigate(loginPage);
    }

    /// <summary>
    /// Strip the back-stack entry every time we navigate. Without this the Frame's
    /// journal grows unbounded across logout/login cycles, holding refs to old pages.
    /// </summary>
    private void RootFrame_Navigated(object sender, NavigationEventArgs e)
    {
        while (RootFrame.CanGoBack) RootFrame.RemoveBackEntry();
    }
}
