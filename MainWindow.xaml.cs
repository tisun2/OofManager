using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using ModernWpf.Controls;
using OofManager.Wpf.Services;
using OofManager.Wpf.Views;

namespace OofManager.Wpf;

public partial class MainWindow : Window
{
    // Preference key: 0 = ask the user every time, 1 = exit, 2 = hide to tray.
    private const string CloseActionPrefKey = "WindowClose.Action";
    private const int ActionAsk = 0;
    private const int ActionExit = 1;
    private const int ActionTray = 2;

    private IPreferencesService? _prefs;
    private ITrayService? _tray;
    // Set to true once the user (or app code) has confirmed an actual exit, so
    // we let Closing through without re-prompting or hiding to tray.
    private bool _exitConfirmed;

    public MainWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Bypass the close-action prompt for the next Close() — used by the tray
    /// "Exit" menu item and the dialog's "Exit" button.
    /// </summary>
    public void ConfirmExit() => _exitConfirmed = true;

    public void Initialize(IServiceProvider services)
    {
        _prefs = (IPreferencesService)services.GetService(typeof(IPreferencesService))!;
        _tray = (ITrayService)services.GetService(typeof(ITrayService))!;

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

    private async void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_exitConfirmed || _prefs == null || _tray == null) return;

        var action = _prefs.GetInt(CloseActionPrefKey, ActionAsk);
        if (action == ActionExit) return;          // let WPF close normally
        if (action == ActionTray)
        {
            e.Cancel = true;
            _tray.HideToTray();
            return;
        }

        // ActionAsk: prompt the user. We must cancel synchronously here — by the
        // time the awaited dialog returns, WPF has already committed to closing.
        e.Cancel = true;
        var (chosen, remember) = await PromptCloseActionAsync();
        if (chosen == ActionTray)
        {
            if (remember) _prefs.Set(CloseActionPrefKey, ActionTray);
            _tray.HideToTray();
        }
        else if (chosen == ActionExit)
        {
            if (remember) _prefs.Set(CloseActionPrefKey, ActionExit);
            _exitConfirmed = true;
            // Re-issue Close on the dispatcher so we exit the cancelled Closing
            // call cleanly before WPF tears down. The discard suppresses CS4014:
            // we deliberately don't await the dispatcher op here.
            _ = Dispatcher.BeginInvoke(new Action(Close));
        }
        // chosen == ActionAsk → user cancelled the dialog; do nothing, window stays open.
    }

    private static async Task<(int action, bool remember)> PromptCloseActionAsync()
    {
        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = "What would you like to do?",
            Margin = new Thickness(0, 0, 0, 12),
            TextWrapping = TextWrapping.Wrap
        });
        var rememberBox = new CheckBox
        {
            Content = "Remember my choice, don't ask again",
            Margin = new Thickness(0, 4, 0, 0)
        };
        stack.Children.Add(rememberBox);

        var dlg = new ContentDialog
        {
            Title = "Close OOF Manager",
            Content = stack,
            PrimaryButtonText = "📥 Hide to Tray",
            SecondaryButtonText = "❌ Exit",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary
        };
        var result = await dlg.ShowAsync();
        var action = result switch
        {
            ContentDialogResult.Primary => ActionTray,
            ContentDialogResult.Secondary => ActionExit,
            _ => ActionAsk
        };
        return (action, rememberBox.IsChecked == true);
    }
}
