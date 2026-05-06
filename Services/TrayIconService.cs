using System.Drawing;
using System.Windows;
using WinForms = System.Windows.Forms;

namespace OofManager.Wpf.Services;

public interface ITrayService
{
    /// <summary>Hide the main window from the taskbar; the tray icon stays visible.</summary>
    void HideToTray();
}

/// <summary>
/// Owns a single Windows.Forms.NotifyIcon for the lifetime of the application.
/// Hide-to-tray is *explicit* (driven by a button bound to <see cref="HideToTray"/>);
/// the standard Minimize button keeps its default behavior of minimizing to the
/// taskbar. Disposed on App.OnExit so we don't leak a ghost icon that lingers in
/// the tray until the user hovers over it.
/// </summary>
public sealed class TrayIconService : ITrayService, IDisposable
{
    private WinForms.NotifyIcon? _icon;
    private Window? _window;
    private readonly IPreferencesService _prefs;
    private bool _disposed;

    public TrayIconService(IPreferencesService prefs)
    {
        _prefs = prefs;
    }

    public void Attach(Window window)
    {
        if (_window != null) return;
        _window = window;

        _icon = new WinForms.NotifyIcon
        {
            Text = "OOF Manager",
            Icon = LoadIcon(),
            Visible = true,
        };

        var menu = new WinForms.ContextMenuStrip();
        menu.Items.Add("Show Window", null, (_, _) => Restore());
        menu.Items.Add("Hide to Tray", null, (_, _) => HideToTray());
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add("Reset Close Button Behavior (Ask Again Next Time)", null, (_, _) =>
        {
            // 0 == ActionAsk in MainWindow. Stored as int; using 0 here keeps
            // the dependency one-way (TrayIconService doesn't import MainWindow).
            _prefs.Set("WindowClose.Action", 0);
        });
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => ExitApp());
        _icon.ContextMenuStrip = menu;

        _icon.MouseClick += (_, e) =>
        {
            // Left click toggles visibility — restore if hidden, hide otherwise.
            // Right-click is handled by the ContextMenuStrip automatically.
            if (e.Button != WinForms.MouseButtons.Left) return;
            if (_window != null && _window.IsVisible) HideToTray();
            else Restore();
        };

        _window.Closed += (_, _) => Dispose();
    }

    public void HideToTray()
    {
        if (_window == null) return;
        // Hide() removes it from the taskbar/Alt-Tab; we explicitly clear
        // ShowInTaskbar too so a quick toggle from another code path doesn't
        // briefly flash a taskbar entry.
        _window.Hide();
        _window.ShowInTaskbar = false;
    }

    private void Restore()
    {
        if (_window == null) return;
        _window.ShowInTaskbar = true;
        _window.Show();
        if (_window.WindowState == WindowState.Minimized)
            _window.WindowState = WindowState.Normal;
        _window.Activate();
        // Topmost flicker is the standard WPF trick to force the window to the
        // foreground when the calling thread doesn't own the active window
        // (which is the case when the click came from the tray's separate
        // message loop). Setting then clearing avoids a permanently-on-top window.
        _window.Topmost = true;
        _window.Topmost = false;
    }

    private static void ExitApp()
    {
        // Bypass the "what should X do?" prompt for the tray-driven exit, then
        // let the standard shutdown path run (which closes the main window and
        // disposes the DI container).
        if (Application.Current?.MainWindow is MainWindow mw)
        {
            mw.ConfirmExit();
        }
        Application.Current?.Shutdown();
    }

    private static Icon LoadIcon()
    {
        try
        {
            // The .ico is embedded as a WPF Resource (see csproj), so reach it
            // via the pack URI rather than a file path — that way it works
            // identically in dev, in publish, and inside the installer.
            var uri = new Uri("pack://application:,,,/Resources/oofmanager.ico", UriKind.Absolute);
            var info = Application.GetResourceStream(uri);
            if (info != null)
            {
                using var s = info.Stream;
                return new Icon(s);
            }
        }
        catch { /* fall through to default */ }
        return SystemIcons.Application;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _window = null;
        if (_icon != null)
        {
            _icon.Visible = false;
            _icon.Dispose();
            _icon = null;
        }
    }
}
