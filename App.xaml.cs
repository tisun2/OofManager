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

        // Local background startup was removed in favour of the Power Automate
        // cloud schedule. Clean up legacy opt-ins so older installs do not keep
        // launching the app at sign-in.
        try
        {
            var prefs = Services.GetRequiredService<IPreferencesService>();
            prefs.Set("WorkSchedule.AutoRefresh", false);
            prefs.Set("WorkSchedule.BackgroundSync", false);
            Services.GetRequiredService<IStartupService>().SetEnabled(false);
        }
        catch { }

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
        var exchange = Services.GetRequiredService<IExchangeService>();
        if (!string.IsNullOrWhiteSpace(lastUpn))
        {
            _ = exchange.TryAutoConnectAsync(lastUpn!, TimeSpan.FromSeconds(45));
        }
        else
        {
            // First-run/manual sign-in path: import ExchangeOnlineManagement while WPF
            // renders so the user's click usually only pays the auth round-trip.
            _ = exchange.PrewarmAsync();
        }

        window.Initialize(Services);
        MainWindow = window;
        window.Show();
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
