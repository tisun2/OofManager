using System.Runtime.InteropServices;

namespace OofManager.Wpf;

/// <summary>
/// Custom entry point. Provides a static helper to lazily allocate the hidden
/// console MSAL/WAM requires — done only the first time we actually need it
/// (right before Connect-ExchangeOnline). Skipping it at startup keeps the
/// app launch flash-free.
/// </summary>
internal static class Program
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AllocConsole();

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    private const int SW_HIDE = 0;
    private const int GWL_STYLE = -16;
    private const int GWL_EXSTYLE = -20;
    private const int WS_VISIBLE = 0x10000000;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_APPWINDOW = 0x00040000;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_HIDEWINDOW = 0x0080;

    private static readonly object _consoleLock = new object();
    private static bool _consolePrepared;

    [STAThread]
    public static void Main()
    {
        // Kick off the heaviest sign-in prep (runspace open + EXO module import)
        // before WPF or DI is touched, so it overlaps with App ctor,
        // InitializeComponent, BuildServiceProvider, MainWindow.Initialize, and
        // first paint. The (singleton) ExchangeService instance later adopts
        // the prepared runspace in PrewarmCoreAsync. Safe and non-blocking.
        try { Services.ExchangeService.BeginEagerPrewarm(); }
        catch { /* best-effort; instance prewarm will still run */ }

        // Returning-user fast path: if a previous successful sign-in remembered a
        // UPN, chain a silent Connect-ExchangeOnline onto the eager runspace right
        // now. By the time WPF, DI, and the LoginPage finish painting, the EXO
        // session is usually already established and ConnectAsync just adopts it
        // instead of paying the ~12s Connect-ExchangeOnline pipeline again.
        try
        {
            var cachedUpn = ReadCachedSignInUpn();
            if (!string.IsNullOrWhiteSpace(cachedUpn))
                Services.ExchangeService.BeginEagerConnect(cachedUpn);
        }
        catch { /* best-effort */ }

        var app = new App();
        app.InitializeComponent();
        app.Run();
    }

    /// <summary>
    /// Reads Auth.LastSignedInUpn from %LOCALAPPDATA%\OofManager\preferences.json
    /// directly, bypassing the DI-resolved PreferencesService so we can use the
    /// value before App.OnStartup runs. Returns null on any error.
    /// </summary>
    private static string? ReadCachedSignInUpn()
    {
        try
        {
            var path = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OofManager",
                "preferences.json");
            if (!System.IO.File.Exists(path)) return null;
            using var doc = System.Text.Json.JsonDocument.Parse(System.IO.File.ReadAllText(path));
            if (doc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Object) return null;
            if (!doc.RootElement.TryGetProperty("Auth.LastSignedInUpn", out var el)) return null;
            if (el.ValueKind != System.Text.Json.JsonValueKind.String) return null;
            return el.GetString();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Idempotently ensures a hidden console window exists for this process.
    /// MSAL/WAM (used by Connect-ExchangeOnline under tenant token-protection)
    /// requires GetConsoleWindow() to return a non-zero HWND. We delay this
    /// allocation until the moment the user actually triggers login, so the
    /// app's startup is completely flash-free; the brief conhost flicker is
    /// then hidden under the WAM auth dialog that appears immediately after.
    /// </summary>
    public static void EnsureHiddenConsole()
    {
        if (_consolePrepared) return;
        lock (_consoleLock)
        {
            if (_consolePrepared) return;
            if (GetConsoleWindow() == IntPtr.Zero)
            {
                if (!AllocConsole())
                {
                    _consolePrepared = true;
                    return;
                }
            }
            var hwnd = GetConsoleWindow();
            if (hwnd != IntPtr.Zero)
            {
                // Strip WS_VISIBLE before it can paint (best-effort).
                var style = GetWindowLong(hwnd, GWL_STYLE);
                SetWindowLong(hwnd, GWL_STYLE, style & ~WS_VISIBLE);
                // Mark as tool window so it never shows on the taskbar / Alt-Tab.
                var ex = GetWindowLong(hwnd, GWL_EXSTYLE);
                ex = (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW;
                SetWindowLong(hwnd, GWL_EXSTYLE, ex);
                // Move off-screen + hide as belt-and-suspenders.
                SetWindowPos(hwnd, IntPtr.Zero, -32000, -32000, 0, 0, SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
                ShowWindow(hwnd, SW_HIDE);
            }
            _consolePrepared = true;
        }
    }
}
