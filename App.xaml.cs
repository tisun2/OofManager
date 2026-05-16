using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using OofManager.Wpf.Services;
using OofManager.Wpf.ViewModels;
using OofManager.Wpf.Views;

namespace OofManager.Wpf;

public partial class App : Application
{
    public IServiceProvider Services { get; private set; } = null!;
    private NavigationService? _navigation;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var window = new MainWindow();

        var services = new ServiceCollection();

        // Services
        services.AddSingleton<IExchangeService, ExchangeService>();
        services.AddSingleton<IPowerAutomateService, PowerAutomateService>();
        services.AddSingleton<ITemplateService, TemplateService>();
        services.AddSingleton<IPreferencesService, PreferencesService>();
        services.AddSingleton<IDialogService, DialogService>();
        services.AddSingleton<IStartupService, StartupService>();
        services.AddSingleton<IWindowsAccountService, WindowsAccountService>();
        // Tray icon: register once, expose under both the concrete type (so
        // OnStartup can call Attach below) and the interface (so ViewModels
        // can request HideToTray without taking a UI dependency).
        services.AddSingleton<TrayIconService>();
        services.AddSingleton<ITrayService>(sp => sp.GetRequiredService<TrayIconService>());
        services.AddSingleton<INavigationService>(sp =>
        {
            _navigation ??= new NavigationService(window.RootFrame, Services);
            return _navigation;
        });

        // ViewModels
        services.AddSingleton<LoginViewModel>();
        services.AddSingleton<MainViewModel>();

        // Views — singletons so navigation is instant (the underlying Frame caches
        // them anyway when JournalEntry is held; making the DI registration explicit
        // means we never rebuild the whole VM/page tree on logout/login).
        services.AddSingleton<LoginPage>();
        services.AddSingleton<MainPage>();

        Services = services.BuildServiceProvider();

        // If the autostart Run entry exists but points at a stale exe (e.g. a
        // previous dev build under bin\Debug, or a prior install location), heal
        // it now so the next user logon launches the actually-installed binary.
        // Cheap; only writes the registry when the path actually differs.
        try { Services.GetRequiredService<IStartupService>().EnsureRegistrationIsFresh(); } catch { }

        // Kick off the PowerShell runspace pre-warm immediately, *before* the window
        // is shown. Opening the runspace and importing ExchangeOnlineManagement takes
        // ~1-2s; running it in parallel with WPF rendering means the user's login
        // click usually finds the module already loaded and only pays the auth time.
        var exchange = Services.GetRequiredService<IExchangeService>();
        _ = exchange.PrewarmAsync();

        // If the user has signed in successfully before, fire a silent reconnect
        // in the background while the WPF window is rendering. Connect-ExchangeOnline
        // with a UPN hint is a no-UI silent token-cache hit when the cache is fresh
        // (~99% of the time once a user has signed in once). By the time the
        // LoginPage is shown, IsConnected is usually already true and the auto-login
        // path navigates straight to MainPage — login feels instant.
        // Only do this when there's a remembered UPN; first-ever launch keeps the
        // explicit "click Sign In" UX so a fresh user never sees a WAM dialog they
        // didn't ask for.
        var lastUpn = Services.GetRequiredService<IPreferencesService>()
            .GetString("Auth.LastSignedInUpn");
        if (!string.IsNullOrWhiteSpace(lastUpn))
        {
            _ = exchange.TryAutoConnectAsync(lastUpn!);
        }

        window.Initialize(Services);
        MainWindow = window;
        window.Show();

        // Attach the tray icon AFTER the window is created. Tray icon is owned
        // by DI so it lives as long as the process and gets disposed exactly
        // once on shutdown (see DisposeAsync above).
        var tray = Services.GetRequiredService<TrayIconService>();
        tray.Attach(window);

        // Honor --minimized: when the OS launched us at user logon (HKCU\Run
        // command line includes the switch) we hide straight to the tray so
        // we don't steal focus from whatever the user opens first.
        if (e.Args.Any(a => string.Equals(a, StartupService.MinimizedArg, StringComparison.OrdinalIgnoreCase)))
        {
            tray.HideToTray();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // Last-chance flush: re-assert the current Scheduled OOF window before
        // the runspace closes. If anything (Outlook desktop, OWA, mobile, admin
        // policy) flipped the mailbox to Disabled while we were running, this
        // is our only opportunity to leave the server in the correct state for
        // the period after OofManager exits. Bounded wait so an EXO hiccup
        // can't hang shutdown — 5s comfortably covers the typical 1-2s
        // Set+Get round-trip without making exit feel stuck.
        //
        // Run on the thread pool (not directly Wait() on the UI thread): the
        // flush awaits Task.Delay / PowerShell invocations that may try to
        // post continuations back to the Dispatcher. A synchronous Wait() on
        // the UI thread would deadlock with those continuations until the 5s
        // timeout elapsed, making shutdown feel hung. Task.Run hops off the
        // Dispatcher so the inner awaits resume on a threadpool thread.
        try
        {
            var vm = Services.GetService<ViewModels.MainViewModel>();
            if (vm != null)
            {
                Task.Run(() => vm.FlushBeforeExitAsync()).Wait(TimeSpan.FromSeconds(5));
            }
        }
        catch { /* best-effort */ }

        // Dispose the DI container on shutdown so singletons that own native handles
        // (PowerShell runspace, SQLite connection) get a chance to flush + close cleanly.
        // Bounded wait so a hung runspace can't hang the whole exit.
        if (Services is IAsyncDisposable asyncDisp)
        {
            try { asyncDisp.DisposeAsync().AsTask().Wait(TimeSpan.FromSeconds(2)); } catch { }
        }
        else if (Services is IDisposable disp)
        {
            try { disp.Dispose(); } catch { }
        }
        base.OnExit(e);
    }
}
