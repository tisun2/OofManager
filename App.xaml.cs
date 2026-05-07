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
        services.AddSingleton<ITemplateService, TemplateService>();
        services.AddSingleton<IPreferencesService, PreferencesService>();
        services.AddSingleton<IDialogService, DialogService>();
        services.AddSingleton<IStartupService, StartupService>();
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
        _ = Services.GetRequiredService<IExchangeService>().PrewarmAsync();

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
