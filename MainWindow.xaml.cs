using System.Windows;
using System.Windows.Navigation;
using OofManager.Wpf.Services;
using OofManager.Wpf.ViewModels;
using OofManager.Wpf.Views;

namespace OofManager.Wpf;

public partial class MainWindow : Window
{
    private ITrayService? _tray;
    private MainViewModel? _vm;
    // Set to true once the user (or app code) has confirmed an actual exit, so
    // we let Closing through without re-routing to hide-to-tray.
    private bool _exitConfirmed;

    public MainWindow()
    {
        InitializeComponent();
        var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        if (v != null) Title = $"OOF Manager v{v.Major}.{v.Minor}.{v.Build}";
    }

    /// <summary>
    /// Bypass the close-to-tray rerouting for the next Close() — used by the
    /// tray "Exit" menu item and any future explicit-exit code paths.
    /// </summary>
    public void ConfirmExit() => _exitConfirmed = true;

    public void Initialize(IServiceProvider services)
    {
        _tray = (ITrayService)services.GetService(typeof(ITrayService))!;
        _vm = (MainViewModel)services.GetService(typeof(MainViewModel))!;

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

        Closing += MainWindow_Closing;
    }

    /// <summary>
    /// Strip the back-stack entry every time we navigate. Without this the Frame's
    /// journal grows unbounded across logout/login cycles, holding refs to old pages.
    /// </summary>
    private void RootFrame_Navigated(object sender, NavigationEventArgs e)
    {
        while (RootFrame.CanGoBack) RootFrame.RemoveBackEntry();
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_exitConfirmed || _vm == null || _tray == null) return;

        // Single rule: if "Run in background" is on, X minimises to tray so
        // the background sync keeps running; if it's off, X really exits.
        // No dialog, no remembered choice — the checkbox itself is the
        // user's persistent preference.
        if (_vm.IsBackgroundSyncEnabled)
        {
            e.Cancel = true;
            _tray.HideToTray();
        }
    }
}
