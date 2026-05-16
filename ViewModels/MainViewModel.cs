using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OofManager.Wpf.Models;
using OofManager.Wpf.Services;

namespace OofManager.Wpf.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly IExchangeService _exchangeService;
    private readonly IPowerAutomateService _powerAutomate;
    private readonly ITemplateService _templateService;
    private readonly IDialogService _dialog;
    private readonly INavigationService _navigation;
    private readonly IPreferencesService _prefs;
    private readonly ITrayService _tray;
    private readonly IStartupService _startup;
    private CancellationTokenSource? _automationCts;
    private bool _hasLoadedOnce;
    // Last (start,end) we actually showed a tray balloon for. Only updates
    // when the *window itself* actually changes, so the 5-min self-heal loop
    // doesn't fire a balloon every tick.
    private DateTimeOffset? _lastNotifiedStart;
    private DateTimeOffset? _lastNotifiedEnd;
    // Last reply text we either fetched from Exchange (LoadAsync) or pushed
    // (any successful SetOofSettingsAsync). The cloud-sync zip / guide
    // generator embeds InternalReply/ExternalReply verbatim, so before
    // shipping the package we compare these snapshots against the live VM
    // properties and prompt if the user has typed unsaved edits — otherwise
    // the cloud flow would silently broadcast a draft that was never
    // confirmed locally.
    private string _committedInternalReply = string.Empty;
    private string _committedExternalReply = string.Empty;
    // Last OOF status pushed to / loaded from the server, kept in lock-step
    // with _committedInternalReply / _committedExternalReply. Drives the
    // "Update reply messages" button's visibility — we only surface it when there's a
    // pending change the user actually needs to push.
    private OofStatus _committedStatus = OofStatus.Disabled;
    // Server-confirmed OOF state for the persistent top status bar. The
    // Manual-mode toggle can now diverge locally until Sync now / Auto-sync
    // pushes it, so the top line must not be derived from IsOofEnabled.
    private OofStatus _confirmedOofStatus = OofStatus.Disabled;
    private DateTimeOffset? _confirmedOofStartTime;
    private DateTimeOffset? _confirmedOofEndTime;
    private bool _isCloudScheduleFlowStatusChecking;
    // Set to true around any programmatic mutation of IsOofEnabled so the
    // partial setter's auto-commit path (which only fires for genuine user
    // gestures on the OOF toggle) doesn't kick a Set against Exchange when
    // we're just reflecting server state locally (LoadAsync, vacation
    // start/end, the auto-sync push, etc.).
    private bool _suppressOofToggleCommit;

    [ObservableProperty] private bool _isBusy;
    // Operation/status feedback for button-driven work (Sync now, package
    // generation, save failures, etc.). This is shown inside the OOF Settings
    // card, not in the persistent top status bar.
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private string _cloudScheduleFlowBannerText = "Power Automate flow: not checked yet.";
    // Persistent top status bar: always describes the real/current OOF state
    // plus the next known OOF window, never transient button progress.
    [ObservableProperty] private string _oofStatusBarMessage = "Loading OOF status...";
    [ObservableProperty] private OofStatus _currentStatus = OofStatus.Disabled;
    [ObservableProperty] private bool _isOofEnabled;
    [ObservableProperty] private bool _isScheduled;
    [ObservableProperty] private bool _isWorkScheduleEnabled;
    /// <summary>
    /// Controls whether the 5-minute background loop keeps re-pushing the
    /// next off-hours window so each day rolls forward automatically. When
    /// off, the schedule is applied once — the moment the user flips the
    /// Work Schedule toggle on, or whenever they save edited boundaries —
    /// and then OofManager stops touching the mailbox until the user takes
    /// another action. Defaults to false: the one-shot push is the
    /// expected behaviour, and opting in to the rolling refresh is a
    /// deliberate choice rather than something that quietly mutates the
    /// mailbox in the background.
    /// </summary>
    [ObservableProperty] private bool _isAutoRefreshEnabled;
    /// <summary>
    /// True when OOF Manager is currently allowed to drive the OOF state for
    /// the user (the Work Schedule rule is on). When false, the OOF on/off
    /// toggle and the reply-update button on the OOF card are interactive again
    /// because the user is in manual control. Kept as a separate property so
    /// the XAML doesn't have to bind directly to <see cref="IsWorkScheduleEnabled"/>
    /// — the two names communicate intent at different cards.
    /// </summary>
    public bool IsOofAutoManaged => IsWorkScheduleEnabled;

    /// <summary>
    /// Mode selector for the segmented control on the Sync card. The two
    /// proxies are kept symmetric so each RadioButton can two-way bind to
    /// its own property without fighting the GroupName mechanism. Both
    /// route through <see cref="IsWorkScheduleEnabled"/> so persistence,
    /// auto-managed gating, and the existing partial-changed handler keep
    /// a single source of truth.
    /// </summary>
    public bool IsScheduleMode
    {
        get => IsWorkScheduleEnabled;
        set { if (value != IsWorkScheduleEnabled) IsWorkScheduleEnabled = value; }
    }
    public bool IsManualMode
    {
        get => !IsWorkScheduleEnabled;
        set { if (value && IsWorkScheduleEnabled) IsWorkScheduleEnabled = false; }
    }

    // Set to true around any programmatic mutation of IsWorkScheduleEnabled
    // (initial prefs load, error rollback) so the partial setter's auto-
    // commit path doesn't kick a fresh Apply against Exchange when we're
    // just reflecting persisted state locally.
    private bool _suppressWorkScheduleCommit;

    /// <summary>
    /// True when the user has edited a reply message since the last successful
    /// save. The OOF switch is intentionally deferred to Sync now / Auto-sync,
    /// so this dirty flag remains reply-text-only for the separate
    /// "Update reply messages" button.
    /// </summary>
    public bool HasUnsavedOofChanges => HasUnsavedReplyChanges();

    public bool HasStatusMessage => !string.IsNullOrWhiteSpace(StatusMessage);

    /// <summary>
    /// The vacation checkbox can only be turned on while OOF itself is on.
    /// Keep it enabled for already-planned/already-active vacations so the
    /// user can still uncheck and sync to clear them.
    /// </summary>
    public bool CanToggleVacationWindow => IsOofEnabled || IsVacationWindowActive || IsOnLongVacation;

    /// <summary>
    /// Caption shown under the Sync card title. Tells the user what the
    /// shared Auto-sync checkbox and ⚡ Sync now button will actually do in
    /// the currently-selected mode, so they don't have to guess whether
    /// "sync" means "push the schedule" or "push my manual state".
    /// </summary>
    public string SyncCardSubtitle
    {
        get
        {
            if (IsScheduleMode)
                return "Schedule mode: OOF Manager auto-flips OOF based on the weekly hours below. ⚡ Sync to Outlook pushes the next off-hours window immediately; Auto-sync re-pushes it every 5 minutes.";
            return "Manual mode: you flip OOF on/off yourself below. ⚡ Sync to Outlook pushes your current state to Outlook immediately; Auto-sync re-pushes it every 5 minutes.";
        }
    }

    public string SyncButtonText
        => IsScheduleMode && HasUnsavedScheduleChanges
            ? "⚡ Save & sync to Outlook"
            : "⚡ Sync to Outlook";

    public string SyncButtonToolTip
    {
        get
        {
            if (IsScheduleMode && HasUnsavedScheduleChanges)
                return "Save your schedule edits, then push the next off-hours window to Outlook.";
            if (IsScheduleMode)
                return "Push the next off-hours window to Outlook right now.";
            return "Push the current OOF on/off state and reply text to Outlook right now.";
        }
    }
    [ObservableProperty] private bool _isStartWithWindowsEnabled;
    // True when any work-schedule field on the panel has been edited since the
    // last successful schedule sync. Drives the shared Sync button's label so
    // the user can tell at a glance when it will save edits first.
    [ObservableProperty] private bool _hasUnsavedScheduleChanges;
    // Suppresses the dirty-flag flip while we're loading prefs from disk or
    // applying a successful save — those property writes are not user edits.
    private bool _suppressDirtyTracking;
    [ObservableProperty] private bool _isMondayWorkday = true;
    [ObservableProperty] private bool _isTuesdayWorkday = true;
    [ObservableProperty] private bool _isWednesdayWorkday = true;
    [ObservableProperty] private bool _isThursdayWorkday = true;
    [ObservableProperty] private bool _isFridayWorkday = true;
    [ObservableProperty] private bool _isSaturdayWorkday = true;
    [ObservableProperty] private bool _isSundayWorkday = true;
    // Per-day work hours. Each day has its own start/end so users can configure
    // e.g. early Mondays and late Fridays. Defaults match the old single-window
    // behaviour (09:00–18:00) until the user persists per-day overrides.
    [ObservableProperty] private TimeSpan _mondayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _mondayEndTime = new(18, 0, 0);
    [ObservableProperty] private TimeSpan _tuesdayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _tuesdayEndTime = new(18, 0, 0);
    [ObservableProperty] private TimeSpan _wednesdayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _wednesdayEndTime = new(18, 0, 0);
    [ObservableProperty] private TimeSpan _thursdayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _thursdayEndTime = new(18, 0, 0);
    [ObservableProperty] private TimeSpan _fridayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _fridayEndTime = new(18, 0, 0);
    [ObservableProperty] private TimeSpan _saturdayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _saturdayEndTime = new(18, 0, 0);
    [ObservableProperty] private TimeSpan _sundayStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _sundayEndTime = new(18, 0, 0);
    [ObservableProperty] private string _workScheduleStatus = "Work schedule rule disabled";
    // Long-vacation mode. Sits "above" the weekly Work Schedule: when on, the
    // automation loop and auto-sync stop fighting Exchange and the OOF window
    // is the single multi-day Scheduled block we pushed in StartVacationAsync.
    // Pre-vacation reply text is squirreled away in prefs so EndVacation can
    // restore the user's normal reply without forcing them to re-type it.
    [ObservableProperty] private bool _isOnLongVacation;
    // User-intent counterpart to IsOnLongVacation. The checkbox in the Manual
    // mode card toggles this; the actual push to Exchange happens later, when
    // the user clicks ⚡ Sync now or the 5-minute auto-sync tick reconciles
    // (see ReassertManualStateAsync). The two stay separate because vacation
    // is a deferred-commit field — flipping the box must not touch the
    // mailbox on its own.
    [ObservableProperty] private bool _isVacationWindowActive;
    [ObservableProperty] private DateTime _vacationStartDate = DateTime.Today;
    [ObservableProperty] private TimeSpan _vacationStartTime = new(18, 0, 0);
    [ObservableProperty] private DateTime _vacationEndDate = DateTime.Today.AddDays(7);
    [ObservableProperty] private TimeSpan _vacationEndTime = new(9, 0, 0);
    [ObservableProperty] private string _vacationStatus = string.Empty;
    [ObservableProperty] private string _internalReply = string.Empty;
    [ObservableProperty] private string _externalReply = string.Empty;
    [ObservableProperty] private DateTime _startDate = DateTime.Today;
    [ObservableProperty] private TimeSpan _startTime = new(9, 0, 0);
    [ObservableProperty] private DateTime _endDate = DateTime.Today.AddDays(1);
    [ObservableProperty] private TimeSpan _endTime = new(9, 0, 0);
    [ObservableProperty] private string _userDisplayName = string.Empty;
    [ObservableProperty] private string _userEmail = string.Empty;
    [ObservableProperty] private string _mailboxIdentity = string.Empty;
    // ObservableCollection so insertions/removals trigger CollectionChanged and the
    // ItemsControl performs an incremental update; assigning a brand-new List<T>
    // forces it to tear down and rebuild every ItemContainer.
    public ObservableCollection<OofTemplate> Templates { get; } = new();
    [ObservableProperty] private OofTemplate? _selectedTemplate;

    // Half-hourly time slots from 00:00 to 23:30 used by the work-time ComboBoxes.
    // Generated once and shared by both the start and end pickers.
    public IReadOnlyList<string> WorkTimeOptions { get; } =
        Enumerable.Range(0, 48)
            .Select(i => TimeSpan.FromMinutes(i * 30).ToString(@"hh\:mm"))
            .ToList();

    public MainViewModel(
        IExchangeService exchangeService,
        IPowerAutomateService powerAutomate,
        ITemplateService templateService,
        IDialogService dialog,
        INavigationService navigation,
        IPreferencesService prefs,
        ITrayService tray,
        IStartupService startup)
    {
        _exchangeService = exchangeService;
        _powerAutomate = powerAutomate;
        _templateService = templateService;
        _dialog = dialog;
        _navigation = navigation;
        _prefs = prefs;
        _tray = tray;
        _startup = startup;
        // Surface the current registry state so the bound CheckBox starts in
        // the right position, even if the user enabled/disabled it externally
        // (Task Manager > Startup, another machine via roaming, etc.).
        _isStartWithWindowsEnabled = _startup.IsEnabled;
    }

    [RelayCommand]
    private void HideToTray() => _tray.HideToTray();

    partial void OnStatusMessageChanged(string value)
        => OnPropertyChanged(nameof(HasStatusMessage));

    partial void OnIsOofEnabledChanged(bool value)
    {
        if (!value)
        {
            CurrentStatus = OofStatus.Disabled;
            IsScheduled = false;
            if (IsVacationWindowActive)
            {
                IsVacationWindowActive = false;
            }
        }
        else
        {
            CurrentStatus = IsScheduled ? OofStatus.Scheduled : OofStatus.Enabled;
        }

        OnPropertyChanged(nameof(CanToggleVacationWindow));
        RefreshOofStatusBar();

        // User-initiated flips are deferred: the switch records local intent
        // only. Exchange is updated by ⚡ Sync now or the 5-minute Auto-sync
        // path (ReassertManualStateAsync). Programmatic sets wrap their
        // writes in _suppressOofToggleCommit so they don't show a pending
        // action note.
        if (_suppressOofToggleCommit) return;
        if (IsOofAutoManaged || !_hasLoadedOnce || IsBusy) return;
        if (IsOnLongVacation)
        {
            StatusMessage = value
                ? "Vacation / Holiday remains selected in Outlook. Click ⚡ Sync to Outlook or enable Auto-sync to re-assert it."
                : "Vacation / Holiday cleared locally. Click ⚡ Sync to Outlook or enable Auto-sync to clear the vacation OOF in Outlook.";
            return;
        }
        StatusMessage = value
            ? "OOF switch set to ON locally. Click ⚡ Sync to Outlook or enable Auto-sync to push it to Outlook."
            : "OOF switch set to OFF locally. Click ⚡ Sync to Outlook or enable Auto-sync to push it to Outlook.";
    }

    partial void OnIsScheduledChanged(bool value)
    {
        if (IsOofEnabled)
        {
            CurrentStatus = value ? OofStatus.Scheduled : OofStatus.Enabled;
        }
        RefreshOofStatusBar();
    }

    partial void OnIsWorkScheduleEnabledChanged(bool value)
    {
        // The Work Schedule toggle is also the master switch for "OofManager
        // is allowed to change OOF state". Refresh every dependent caption
        // and binding so the OOF card flips between interactive and read-
        // only mode immediately.
        WorkScheduleStatus = value ? GetWorkScheduleStatus() : "Work schedule rule disabled";
        OnPropertyChanged(nameof(IsOofAutoManaged));
        OnPropertyChanged(nameof(IsScheduleMode));
        OnPropertyChanged(nameof(IsManualMode));
        OnPropertyChanged(nameof(SyncCardSubtitle));
        OnPropertyChanged(nameof(SyncButtonText));
        OnPropertyChanged(nameof(SyncButtonToolTip));
        RefreshOofStatusBar();

        if (!value && !_suppressDirtyTracking && !_suppressWorkScheduleCommit)
        {
            _ = RefreshCloudScheduleFlowStatusAsync();
        }

        // Initial hydration (LoadWorkSchedulePreferences) and error rollbacks
        // pass through the same setter; we only commit for genuine user
        // flips. _suppressDirtyTracking is set during prefs load, and the
        // dedicated _suppressWorkScheduleCommit flag is set by rollbacks.
        if (_suppressDirtyTracking || _suppressWorkScheduleCommit) return;
        _ = CommitWorkScheduleFlipAsync(value);
    }

    partial void OnIsAutoRefreshEnabledChanged(bool value)
    {
        // Merged "Run in background" checkbox is computed from both this and
        // IsStartWithWindowsEnabled; fire before the suppress-return so the UI
        // also reflects state loaded from prefs during initial hydration.
        OnPropertyChanged(nameof(IsBackgroundSyncEnabled));
        // No mailbox push needed — this flag only gates the in-loop work.
        // Persist the choice so it survives restarts. The loop itself reads
        // IsAutoRefreshEnabled on every tick, so flipping this checkbox
        // takes effect on the very next tick without restarting the loop.
        if (_suppressDirtyTracking) return;
        _prefs.Set("WorkSchedule.AutoRefresh", value);
        OnPropertyChanged(nameof(SyncCardSubtitle));

        // Manual mode + Auto-sync ON needs the loop running for periodic
        // reconcile against Exchange. StartAutomationLoop is idempotent
        // (no-ops when the loop is already running), so calling it here
        // unconditionally on ON-flips is safe regardless of mode.
        if (value && _hasLoadedOnce && !IsOnLongVacation)
        {
            StartAutomationLoop();
        }
    }

    partial void OnCurrentStatusChanged(OofStatus value)
    {
        RefreshOofStatusBar();
    }

    partial void OnInternalReplyChanged(string value)
    {
    }

    partial void OnExternalReplyChanged(string value)
    {
    }
    // Vacation toggle drives both the work-schedule status caption (so the
    // user can see "paused" right on the schedule card) and a refresh of the
    // schedule-status string itself.
    partial void OnIsOnLongVacationChanged(bool value)
    {
        WorkScheduleStatus = GetWorkScheduleStatus();
        OnPropertyChanged(nameof(CanToggleVacationWindow));
        if (value) EnsureVacationShowsManualOofOn();
        RefreshOofStatusBar();
    }

    // Persist the checkbox state so a restart doesn't lose a planned
    // vacation that hasn't been synced yet. The actual push to Exchange is
    // deferred until the user clicks ⚡ Sync now or the 5-minute auto-sync
    // tick reconciles via ReassertManualStateAsync.
    partial void OnIsVacationWindowActiveChanged(bool value)
    {
        if (value && !IsOofEnabled && !IsOnLongVacation)
        {
            IsVacationWindowActive = false;
            return;
        }

        if (value) EnsureVacationShowsManualOofOn();

        OnPropertyChanged(nameof(CanToggleVacationWindow));
        RefreshOofStatusBar();

        if (_suppressDirtyTracking) return;
        _prefs.Set("Vacation.WindowActive", value);

        // Default the start to "now floored to the previous :00 or :30" so
        // the user isn't scrolling back from a hard-coded 18:00 every time.
        // Only fires on a genuine user gesture — prefs-restore returned above.
        if (value)
        {
            var now = DateTime.Now;
            var floor = new DateTime(now.Year, now.Month, now.Day, now.Hour, (now.Minute / 30) * 30, 0);
            VacationStartDate = floor.Date;
            VacationStartTime = floor.TimeOfDay;
        }
    }
    partial void OnStartDateChanged(DateTime value) => RefreshOofStatusBar();
    partial void OnStartTimeChanged(TimeSpan value) => RefreshOofStatusBar();
    partial void OnEndDateChanged(DateTime value) => RefreshOofStatusBar();
    partial void OnEndTimeChanged(TimeSpan value) => RefreshOofStatusBar();
    partial void OnVacationStartDateChanged(DateTime value) => RefreshOofStatusBar();
    partial void OnVacationStartTimeChanged(TimeSpan value) => RefreshOofStatusBar();
    partial void OnVacationEndDateChanged(DateTime value) => RefreshOofStatusBar();
    partial void OnVacationEndTimeChanged(TimeSpan value) => RefreshOofStatusBar();

    private void EnsureVacationShowsManualOofOn()
    {
        if ((!IsVacationWindowActive && !IsOnLongVacation) || IsOofEnabled) return;
        _suppressOofToggleCommit = true;
        try { IsOofEnabled = true; } finally { _suppressOofToggleCommit = false; }
    }

    partial void OnIsMondayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnIsTuesdayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnIsWednesdayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnIsThursdayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnIsFridayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnIsSaturdayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnIsSundayWorkdayChanged(bool value) => MarkScheduleDirty();
    partial void OnMondayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnMondayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnTuesdayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnTuesdayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnWednesdayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnWednesdayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnThursdayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnThursdayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnFridayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnFridayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnSaturdayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnSaturdayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnSundayStartTimeChanged(TimeSpan value) => MarkScheduleDirty();
    partial void OnSundayEndTimeChanged(TimeSpan value) => MarkScheduleDirty();

    /// <summary>
    /// Flips HasUnsavedScheduleChanges true unless we're currently loading from
    /// disk or applying our own post-save cleanup. The Save button binds to
    /// this so users can tell at a glance whether their edits are persisted.
    /// </summary>
    private void MarkScheduleDirty()
    {
        if (_suppressDirtyTracking) return;
        HasUnsavedScheduleChanges = true;
        RefreshOofStatusBar();
    }

    partial void OnHasUnsavedScheduleChangesChanged(bool value)
    {
        OnPropertyChanged(nameof(SyncButtonText));
        OnPropertyChanged(nameof(SyncButtonToolTip));
    }

    partial void OnIsStartWithWindowsEnabledChanged(bool value)
    {
        // Two-way bind directly to the registry. We re-read after writing so a
        // failed write (locked-down policy) flips the bound CheckBox back to
        // its real state instead of silently lying to the user.
        _startup.SetEnabled(value);
        var actual = _startup.IsEnabled;
        if (actual != value)
        {
            // Re-set without re-entering this handler: assigning the field
            // directly skips the generated setter's change-detection.
            _isStartWithWindowsEnabled = actual;
            OnPropertyChanged(nameof(IsStartWithWindowsEnabled));
            StatusMessage = "Could not update startup setting (policy may block it).";
        }
        // Merged "Run in background" checkbox reads from both this and
        // IsAutoRefreshEnabled; fire after the policy-rollback above so the
        // UI converges on the *actual* registry state.
        OnPropertyChanged(nameof(IsBackgroundSyncEnabled));
    }

    /// <summary>
    /// Merged user-facing background-mode toggle. Bundles the two flags that
    /// only make sense together:
    ///   • <see cref="IsStartWithWindowsEnabled"/> — auto-launch at logon, so
    ///     the app is actually running to do anything in the background.
    ///   • <see cref="IsAutoRefreshEnabled"/> — the in-loop 5-minute reconcile.
    /// Reads ON when *either* underlying flag is on, so a user who upgrades
    /// from a build that had them separate doesn't see the merged checkbox
    /// lie about state. Writing flips both at once.
    /// </summary>
    public bool IsBackgroundSyncEnabled
    {
        get => IsStartWithWindowsEnabled || IsAutoRefreshEnabled;
        set
        {
            IsStartWithWindowsEnabled = value;
            IsAutoRefreshEnabled = value;
        }
    }

    partial void OnSelectedTemplateChanged(OofTemplate? value)
    {
        if (value == null) return;
        ApplyTemplate(value);
    }

    /// <summary>
    /// Loads a template's content into the reply editors.
    /// Bound directly to the "Load" button so it always fires, even if the
    /// user clicks the same template twice in a row (which would otherwise
    /// be a no-op for the SelectedTemplate setter).
    /// </summary>
    [RelayCommand]
    private void LoadTemplate(OofTemplate? template)
    {
        if (template == null) return;
        ApplyTemplate(template);
    }

    private void ApplyTemplate(OofTemplate template)
    {
        InternalReply = template.InternalReply;
        ExternalReply = template.ExternalReply;
        StatusMessage = $"Loaded template \u201c{template.Name}\u201d";
    }

    [RelayCommand]
    public async Task LoadAsync()
    {
        if (IsBusy) return;
        // Page.Loaded fires every time the user navigates back. Only do the
        // full server fetch once; subsequent activations are instant.
        if (_hasLoadedOnce) return;

        IsBusy = true;
        StatusMessage = "Loading OOF settings...";

        // MainPage may have been navigated to directly at startup (when a
        // remembered UPN exists) before Connect-ExchangeOnline has finished.
        // Wait for the in-flight silent reconnect; if there is no IsConnected
        // by the time it returns, the silent attempt failed — fall back to the
        // LoginPage so the user can sign in interactively.
        if (!_exchangeService.IsConnected)
        {
            var lastUpn = _prefs.GetString("Auth.LastSignedInUpn");
            if (!string.IsNullOrWhiteSpace(lastUpn))
            {
                await _exchangeService.TryAutoConnectAsync(lastUpn!);
            }
            if (!_exchangeService.IsConnected)
            {
                IsBusy = false;
                _navigation.NavigateToLogin();
                return;
            }
        }

        try
        {
            // Kick off the slow network fetch and the local DB read in parallel
            // immediately — don't gate them on GetCurrentUserAsync (which is just
            // a cached field read but conceptually a precondition).
            var oofTask = _exchangeService.GetOofSettingsAsync();
            var templatesTask = _templateService.GetAllTemplatesAsync();

            var signedInUser = await _exchangeService.GetCurrentUserAsync();
            var mailboxIdentity = await _exchangeService.GetCurrentMailboxIdentityAsync();
            MailboxIdentity = mailboxIdentity;
            UserDisplayName = mailboxIdentity;
            UserEmail = string.Equals(mailboxIdentity, signedInUser, StringComparison.OrdinalIgnoreCase)
                ? signedInUser
                : $"Signed in as {signedInUser}";
            LoadWorkSchedulePreferences();

            await Task.WhenAll(oofTask, templatesTask);

            var oof = oofTask.Result;
            // The toggle should reflect what Outlook is *actually doing right
            // now*, not the raw mailbox state. A Scheduled OOF whose window is
            // still in the future means OOF replies aren't being sent yet, so
            // we don't want the toggle to read "On". Without this guard, the
            // auto-sync's "next off-hours window" pre-push would flip the
            // toggle on as soon as the user signs in, even mid-workday.
            // The suppress flag stops the toggle-flip auto-commit path from
            // echoing this server state back to Exchange as a fresh Set.
            _suppressOofToggleCommit = true;
            try
            {
                IsScheduled = oof.Status == OofStatus.Scheduled;
                IsOofEnabled = oof.Status == OofStatus.Enabled
                    || (oof.Status == OofStatus.Scheduled
                        && oof.StartTime.HasValue
                        && oof.StartTime.Value <= DateTimeOffset.Now
                        && (!oof.EndTime.HasValue || oof.EndTime.Value > DateTimeOffset.Now));
            }
            finally
            {
                _suppressOofToggleCommit = false;
            }
            // Set CurrentStatus *last* and unconditionally to the real mailbox
            // state. The partial-method side effects of IsOofEnabled /
            // IsScheduled assignments above flip CurrentStatus around as if
            // the user had toggled the switch, which would otherwise lie to
            // the status label about what Exchange actually has.
            CurrentStatus = oof.Status;
            InternalReply = oof.InternalReply;
            ExternalReply = oof.ExternalReply;
            MarkOofClean(oof.Status);
            RememberConfirmedOofState(oof);

            if (oof.StartTime.HasValue)
            {
                StartDate = oof.StartTime.Value.LocalDateTime.Date;
                StartTime = oof.StartTime.Value.LocalDateTime.TimeOfDay;
            }
            if (oof.EndTime.HasValue)
            {
                EndDate = oof.EndTime.Value.LocalDateTime.Date;
                EndTime = oof.EndTime.Value.LocalDateTime.TimeOfDay;
            }
            EnsureVacationShowsManualOofOn();

            Templates.Clear();
            foreach (var t in templatesTask.Result) Templates.Add(t);

            _hasLoadedOnce = true;
            RefreshOofStatusBar();
            StatusMessage = string.Empty;

            if (IsOnLongVacation)
            {
                // Vacation is the source of truth right now; don't let the
                // work-schedule code path run a sync that would overwrite the
                // long Scheduled window. Still start the automation loop so
                // the loop's HasVacationEnded check can auto-clear at the
                // configured end time, even if Work Schedule is off.
                IsBusy = false;
                StartAutomationLoop();

                // Also self-heal: if the vacation end time is already in the
                // past (machine off when vacation expired, etc.), clean up
                // immediately so the user doesn't see a stale "On vacation"
                // banner that won't go away until the next 5-min tick.
                if (HasVacationEnded(DateTime.Now))
                {
                    await EndVacationAsync();
                }
            }
            else if (IsWorkScheduleEnabled)
            {
                // Drop the busy flag *before* refreshing the schedule caption.
                // We deliberately do NOT push to Outlook on launch: the toggle
                // is the "OofManager is in charge" gate, but the mailbox push
                // is opt-in (Auto-refresh ticker or ⚡ Sync now). Pushing
                // unconditionally on every launch would mutate Outlook out
                // from under users who never asked for it.
                IsBusy = false;
                StartAutomationLoop();
                await ApplyWorkScheduleAsync(showSuccessMessage: false, suppressTrayNotification: true);
            }
            else if (IsAutoRefreshEnabled)
            {
                // Manual mode + Auto-sync persisted as ON: spin up the loop so
                // the 5-min reconcile against Exchange resumes after restart.
                IsBusy = false;
                StartAutomationLoop();
            }
        }
        catch (Exception ex)
        {
            OofStatusBarMessage = "OOF status unavailable";
            StatusMessage = $"Load failed: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            if (_hasLoadedOnce && IsManualMode)
            {
                _ = RefreshCloudScheduleFlowStatusAsync();
            }
        }
    }

    /// <summary>
    /// Auto-commit path for the Work Schedule toggle. Mirrors the OOF card
    /// toggle's behaviour: flipping the switch sends the result straight
    /// to Exchange, so the user no longer has to chase a separate "Apply"
    /// button. Saves the in-panel schedule prefs as a side effect when
    /// turning on so anything the user edited in the grid takes effect
    /// immediately. On failure the toggle reverts so it doesn't lie about
    /// the actual server state.
    /// </summary>
    private async Task CommitWorkScheduleFlipAsync(bool turnOn)
    {
        if (turnOn)
        {
            // Validate grid before persisting / pushing. If anything's bad,
            // pop the toggle back off without writing to Exchange.
            foreach (var day in WeekDays)
            {
                if (!IsWorkday(day)) continue;
                if (GetEndTimeForDay(day) <= GetStartTimeForDay(day))
                {
                    _suppressWorkScheduleCommit = true;
                    try { IsWorkScheduleEnabled = false; }
                    finally { _suppressWorkScheduleCommit = false; }
                    await _dialog.AlertAsync(
                        "Invalid Work Hours",
                        $"{GetDayDisplayName(day)}: end time must be later than start time.");
                    return;
                }
            }
        }

        SaveWorkSchedulePreferences();
        HasUnsavedScheduleChanges = false;

        if (turnOn)
        {
            StartAutomationLoop();
            // Mode changes are local UI/model changes. The OOF Settings subtitle
            // already explains what Sync to Outlook and Auto-sync will do, so
            // don't echo another status line underneath it.
            StatusMessage = string.Empty;
            await Task.CompletedTask;
        }
        else
        {
            // Keep the loop alive when Auto-sync is on — Manual mode also
            // needs the 5-min ticker for periodic reconcile against Exchange.
            // Otherwise the loop is idle work, so shut it down.
            if (!IsAutoRefreshEnabled)
            {
                StopAutomationLoop();
            }
            WorkScheduleStatus = "Work schedule rule disabled";

            // Switching to Manual mode is now a local mode change only. Do
            // not collapse an existing Scheduled window to flat Enabled /
            // Disabled here; that was the path that made the top status jump
            // to "OOF is ON — no end time" before the user clicked Sync now.
            StatusMessage = string.Empty;
        }
    }

    [RelayCommand]
    private async Task SaveWorkScheduleAsync()
    {
        if (!await TrySavePendingScheduleChangesAsync()) return;

        if (IsWorkScheduleEnabled)
        {
            await ApplyWorkScheduleAsync(showSuccessMessage: true, suppressTrayNotification: true);
            await SyncToOutlookCoreAsync(isUserInitiated: false);
        }
        else
        {
            // Same loop-lifetime rule as the toggle path: keep the ticker
            // alive for Manual-mode reconcile when Auto-sync is on.
            if (!IsAutoRefreshEnabled)
            {
                StopAutomationLoop();
            }
            StatusMessage = "Work schedule rule saved but not enabled";
            WorkScheduleStatus = "Work schedule rule disabled";
        }
    }

    /// <summary>
    /// "⚡ Sync now" button. The semantics depend on the current sync mode:
    ///
    ///   • Schedule mode: re-evaluates "are we inside working hours?" and
    ///     force-pushes the next off-hours window to Outlook (formerly the
    ///     only behaviour of this button).
    ///   • Manual mode: re-asserts the user's current on/off + reply text
    ///     to Exchange immediately, collapsing any leftover Scheduled state
    ///     from a previous schedule-mode session. Useful when another
    ///     client (Outlook desktop, OWA, an admin) drifted the mailbox
    ///     away from what this app shows.
    /// </summary>
    [RelayCommand]
    private async Task SyncNowAsync()
    {
        if (IsScheduleMode)
        {
            if (!await TrySavePendingScheduleChangesAsync()) return;

            if (IsOnLongVacation)
            {
                StatusMessage = "Clearing Vacation / Holiday override and syncing the schedule...";
                await EndVacationAsync();
                if (IsOnLongVacation) return;
                StatusMessage = "Schedule mode synced. Vacation / Holiday override cleared.";
                return;
            }

            if (IsVacationWindowActive)
            {
                IsVacationWindowActive = false;
                VacationStatus = string.Empty;
            }

            // 1. Local re-check + Exchange OOF flip.
            await ApplyWorkScheduleAsync(showSuccessMessage: false, suppressTrayNotification: true);

            // 2. Force-push the next window to Outlook. SyncToOutlookCoreAsync
            //    skips its server-state dedupe when isUserInitiated=true, so
            //    this always re-asserts the desired state.
            await SyncToOutlookCoreAsync(isUserInitiated: true);
            return;
        }

        // Manual mode: re-push current local state straight to Exchange.
        await ReassertManualStateAsync(isUserInitiated: true);
    }

    private async Task<bool> TrySavePendingScheduleChangesAsync()
    {
        if (!HasUnsavedScheduleChanges) return true;

        // Validate every enabled day individually — disabled days don't participate
        // in the automation loop, so their time values are ignored.
        foreach (var day in WeekDays)
        {
            if (!IsWorkday(day)) continue;
            var start = GetStartTimeForDay(day);
            var end = GetEndTimeForDay(day);
            if (end <= start)
            {
                await _dialog.AlertAsync(
                    "Invalid Work Hours",
                    $"{GetDayDisplayName(day)}: end time must be later than start time.");
                return false;
            }
        }

        SaveWorkSchedulePreferences();
        HasUnsavedScheduleChanges = false;
        return true;
    }

    /// <summary>
    /// Pushes the current Manual-mode "desired state" (the OOF on/off toggle
    /// + reply text) to Exchange. Shared by the user-pressed ⚡ Sync now
    /// button and the 5-minute background auto-sync; the latter sets
    /// <paramref name="isUserInitiated"/>=false so the call reads server
    /// state first and skips the Set when nothing actually drifted.
    /// </summary>
    private async Task ReassertManualStateAsync(bool isUserInitiated)
    {
        if (IsBusy) return;

        // Vacation reconciliation runs first — the checkbox is a deferred-
        // commit field, so a divergence between IsVacationWindowActive and
        // IsOnLongVacation means this sync should push the start/end. After
        // either branch fires we return: StartVacationAsync/EndVacationAsync
        // both end up writing to Exchange themselves, and OOF on/off must
        // not run on top.
        if (IsVacationWindowActive && !IsOnLongVacation)
        {
            await StartVacationAsync();
            return;
        }
        if (!IsVacationWindowActive && IsOnLongVacation)
        {
            await EndVacationAsync();
            return;
        }
        // Vacation owns the OOF window outright; never let manual reassert
        // overwrite the multi-day Scheduled block.
        if (IsOnLongVacation) return;

        var desired = new OofSettings
        {
            Status = IsOofEnabled ? OofStatus.Enabled : OofStatus.Disabled,
            InternalReply = InternalReply,
            ExternalReply = ExternalReply,
        };

        // Auto path: read live server state and skip the Set when the
        // mailbox already matches. The user path always re-asserts.
        if (!isUserInitiated)
        {
            try
            {
                var current = await _exchangeService.GetOofSettingsAsync();
                if (current.Status == desired.Status
                    && string.Equals(current.InternalReply ?? string.Empty, desired.InternalReply ?? string.Empty, StringComparison.Ordinal)
                    && string.Equals(current.ExternalReply ?? string.Empty, desired.ExternalReply ?? string.Empty, StringComparison.Ordinal))
                {
                    RememberConfirmedOofState(current);
                    return;
                }
            }
            catch
            {
                // Couldn't verify (network blip etc.) — fall through and Set.
                // SetOofSettingsAsync verifies on its own, so we still won't
                // silently succeed on a real failure.
            }
        }

        IsBusy = true;
        try
        {
            StatusMessage = isUserInitiated
                ? "📤 Re-asserting OOF state to Outlook..."
                : "🔄 Auto-syncing OOF state...";
            await _exchangeService.SetOofSettingsAsync(desired);
            RememberConfirmedOofState(desired);

            // Collapse any lingering Scheduled flag so the UI matches what
            // we just pushed (flat Enabled or Disabled).
            _suppressOofToggleCommit = true;
            try { IsScheduled = false; } finally { _suppressOofToggleCommit = false; }
            CurrentStatus = desired.Status;
            MarkOofClean(desired.Status);

            if (isUserInitiated)
            {
                StatusMessage = desired.Status == OofStatus.Enabled
                    ? "✅ OOF re-asserted as ON in Outlook."
                    : "✅ OOF re-asserted as OFF in Outlook.";
            }
        }
        catch (Exception ex)
        {
            if (isUserInitiated)
            {
                StatusMessage = $"Sync failed: {ex.Message}";
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Generates a personalised Power Automate setup guide (HTML) that the
    /// user can follow to create a cloud flow which keeps their OOF window
    /// in sync even when every local computer is off. Bound to the
    /// "Generate cloud schedule setup guide" button in the Work Schedule card.
    /// </summary>
    [RelayCommand]
    private async Task OpenCloudScheduleGuideAsync()
    {
        try
        {
            var snapshot = new WorkScheduleSnapshot(_prefs);
            var path = CloudScheduleGuideGenerator.GenerateAndOpen(
                snapshot,
                userEmail: MailboxIdentity,
                internalReply: InternalReply,
                externalReply: ExternalReply,
                externalAudienceAll: true);
            StatusMessage = $"🌐 Cloud schedule setup guide opened in your browser ({System.IO.Path.GetFileName(path)})";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to generate cloud schedule guide: {ex.Message}";
            await _dialog.AlertAsync("Cloud Schedule Guide", ex.Message);
        }
    }

    /// <summary>
    /// Builds the OofManager Cloud Schedule solution package on the user's
    /// Desktop and opens make.powerautomate.com/solutions in their browser
    /// so they can drop the zip into the Import dialog. We previously
    /// auto-imported via Dataverse Web API, but locked-down tenants (e.g.
    /// Microsoft corp) block the bundled PowerApps PowerShell module's
    /// client ID from minting Dataverse tokens (AADSTS65002), so the
    /// reliable path is to hand the zip to the user and let the maker
    /// portal's own import dialog do the upload.
    /// </summary>
    [RelayCommand]
    private async Task GenerateCloudSchedulePackageAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        try
        {
            var snapshot = new WorkScheduleSnapshot(_prefs);
            // Drop the zip on the Desktop so it's easy to find from the
            // Power Automate import file picker.
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var outPath = System.IO.Path.Combine(desktop, "OofManager-CloudSchedule.zip");

            StatusMessage = "📦 Generating cloud schedule solution…";
            var pkg = await Task.Run(() => CloudSchedulePackageGenerator.GenerateWithIdentity(
                snapshot,
                userEmail: MailboxIdentity,
                internalReply: InternalReply,
                externalReply: ExternalReply,
                externalAudienceAll: true,
                generateManaged: false,
                outputPath: outPath));

            OpenSolutionsPage();
            StatusMessage = "✅ Saved OofManager-CloudSchedule.zip to your Desktop. Power Automate opened — switch to the environment named after you, click 'Import solution', and pick the zip from your Desktop.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to generate cloud schedule solution: {ex.Message}";
            await _dialog.AlertAsync("Cloud Schedule Solution", ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Opens the Solutions page in the user's browser. When the env id is
    /// known we deep-link to that env's solutions page so the user doesn't
    /// have to pick from the env switcher.
    /// </summary>
    private static void OpenSolutionsPage(string? environmentId = null)
    {
        var url = !string.IsNullOrWhiteSpace(environmentId)
            ? $"https://make.powerautomate.com/environments/{environmentId}/solutions"
            : "https://make.powerautomate.com/solutions";
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url)
            {
                UseShellExecute = true,
            });
        }
        catch { /* best-effort */ }
    }

    /// <summary>
    /// Opens make.powerautomate.com/flows so the user can locate the imported
    /// OofManager Cloud Schedule flow and toggle it Off. Lightweight companion to
    /// the Manual-mode "⚠️ Cloud Schedule flow runs in M365" reminder banner —
    /// saves the user from copy-pasting the URL into a browser by hand. We
    /// don't pin an environment or flow id in the URL: the user could be in
    /// any environment, and the plain /flows page lands them in whichever env
    /// the portal had them in last (env picker top-right), where the imported
    /// flow shows up under <em>Cloud flows</em> with a one-click On/Off toggle.
    /// </summary>
    [RelayCommand]
    private void OpenPowerAutomateFlows()
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                "https://make.powerautomate.com/flows")
            {
                UseShellExecute = true,
            });
        }
        catch
        {
            // Best-effort: if the OS has no default browser handler we just
            // swallow — the banner text still tells the user where to go.
        }
    }

    /// <summary>
    /// One-click "turn off the cloud schedule flow" — calls the bundled
    /// Microsoft.PowerApps.PowerShell module via a hidden child process and
    /// flips the flow's Off switch the same way the Power Automate UI does.
    /// First invocation pops the module's auth dialog; subsequent runs are
    /// silent because the module caches its own refresh token. Any
    /// non-Success outcome falls back to opening make.powerautomate.com so
    /// the user can finish the job manually.
    /// </summary>
    [RelayCommand]
    private Task DisableCloudScheduleFlowAsync() => RunCloudScheduleFlowToggleAsync(disable: true);

    /// <summary>
    /// Symmetric counterpart to <see cref="DisableCloudScheduleFlowAsync"/>.
    /// Re-enables the flow after the user toggled it off (manual mode banner
    /// button) and wants the cloud schedule running again.
    /// </summary>
    [RelayCommand]
    private Task EnableCloudScheduleFlowAsync() => RunCloudScheduleFlowToggleAsync(disable: false);

    private async Task RefreshCloudScheduleFlowStatusAsync()
    {
        if (!IsManualMode || _isCloudScheduleFlowStatusChecking) return;

        _isCloudScheduleFlowStatusChecking = true;
        CloudScheduleFlowBannerText = "Power Automate flow: checking...";
        try
        {
            var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
            var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
            var expectedFlowDisplayName = CloudSchedulePackageGenerator.ComputeFlowIdentity(upn ?? string.Empty).FlowDisplayName;
            var result = await _powerAutomate.GetOofManagerFlowStatusAsync(upn, displayName, expectedFlowDisplayName);
            if (IsManualMode)
            {
                SetCloudScheduleFlowBanner(result.State);
            }
        }
        catch
        {
            if (IsManualMode)
            {
                SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
            }
        }
        finally
        {
            _isCloudScheduleFlowStatusChecking = false;
        }
    }

    private void SetCloudScheduleFlowBanner(PowerAutomateFlowState state)
    {
        CloudScheduleFlowBannerText = state switch
        {
            PowerAutomateFlowState.On => "Power Automate flow: On. Turn it off before vacation; turn it back on after.",
            PowerAutomateFlowState.Off => "Power Automate flow: Off. Turn it back on after vacation to resume the cloud schedule.",
            PowerAutomateFlowState.NotFound => "Power Automate flow: Not found. Import the Cloud Schedule package first.",
            _ => "Power Automate flow: Unknown. Sign in to Power Automate to check or use the buttons.",
        };
    }

    private async Task RunCloudScheduleFlowToggleAsync(bool disable)
    {
        if (IsBusy) return;
        var verbLabel = disable ? "Turning off" : "Turning on";
        IsBusy = true;
        StatusMessage = $"{verbLabel} Power Automate flow… using cached Power Automate sign-in if available. First run may show a sign-in dialog; otherwise this runs quietly and can take a few seconds.";
        try
        {
            var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
            // Display name is used to prioritise environments named after the
            // user (e.g. 'Sandy Sun's Environment') before falling back to
            // Default / capped scan — Microsoft-style tenants typically have
            // hundreds of envs the signed-in account is merely visible to.
            var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
            // Recompute the per-user flow display name locally — same formula
            // the package generator uses at import time — so the PS script can
            // match Get-Flow's DisplayName exactly (e.g. 'OofManager Cloud Schedule
            // (TianyueSun)') instead of relying on a prefix that could collide
            // with unrelated flows. Power Automate doesn't preserve the
            // workflow GUID we stamp into solution.xml at import, so the
            // display-name suffix is the most reliable per-user key.
            var expectedFlowDisplayName = CloudSchedulePackageGenerator.ComputeFlowIdentity(upn ?? string.Empty).FlowDisplayName;
            var result = disable
                ? await _powerAutomate.DisableOofManagerFlowsAsync(upn, displayName, expectedFlowDisplayName)
                : await _powerAutomate.EnableOofManagerFlowsAsync(upn, displayName, expectedFlowDisplayName);
            var flowDisplayNames = result.FlowDisplayNames.Count > 0
                ? string.Join(", ", result.FlowDisplayNames)
                : expectedFlowDisplayName;

            switch (result.Outcome)
            {
                case PowerAutomateOutcome.Success:
                    SetCloudScheduleFlowBanner(disable ? PowerAutomateFlowState.Off : PowerAutomateFlowState.On);
                    StatusMessage = string.IsNullOrWhiteSpace(flowDisplayNames)
                        ? "✅ " + result.Message
                        : $"✅ {result.Message}: {flowDisplayNames}";
                    break;
                case PowerAutomateOutcome.NoFlowFound:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.NotFound);
                    StatusMessage = "⚠️ " + result.Message;
                    OpenPowerAutomateFlows();
                    break;
                case PowerAutomateOutcome.SignInFailed:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
                    StatusMessage = "⚠️ Couldn't sign in to Power Automate. Opening the website so you can toggle the flow manually.";
                    OpenPowerAutomateFlows();
                    break;
                case PowerAutomateOutcome.SolutionAwareBlocked:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
                    StatusMessage = "⚠️ Power Automate refused the toggle (solution-aware flow). Opening the website so you can use the Off switch directly.";
                    OpenPowerAutomateFlows();
                    break;
                default:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
                    StatusMessage = "⚠️ " + result.Message + " — opening Power Automate so you can finish manually.";
                    OpenPowerAutomateFlows();
                    break;
            }
        }
        catch (Exception ex)
        {
            SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
            StatusMessage = $"Failed to toggle Power Automate flow: {ex.Message}";
            OpenPowerAutomateFlows();
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool HasUnsavedReplyChanges()
    {
        return !string.Equals(InternalReply ?? string.Empty, _committedInternalReply, StringComparison.Ordinal)
            || !string.Equals(ExternalReply ?? string.Empty, _committedExternalReply, StringComparison.Ordinal);
    }

    /// <summary>
    /// Snapshots the current OOF state (status + reply bodies) as the
    /// committed baseline. Called after any successful Set against Exchange.
    /// </summary>
    private void MarkOofClean(OofStatus committedStatus)
    {
        _committedInternalReply = InternalReply ?? string.Empty;
        _committedExternalReply = ExternalReply ?? string.Empty;
        _committedStatus = committedStatus;
    }

    /// <summary>
    /// Starts a long-vacation block: pushes a single multi-day Scheduled OOF
    /// window to Exchange covering the entire vacation, stashes the user's
    /// current reply text in prefs (so EndVacation can restore it), and pauses
    /// the regular Work Schedule auto-toggle/auto-sync so it doesn't fight the
    /// vacation window. To use a saved reply, the user picks one from the
    /// Templates card first; whatever is in the Internal/External boxes at
    /// click time is what we send.
    /// </summary>
    [RelayCommand]
    private async Task StartVacationAsync()
    {
        if (IsBusy) return;
        var startLocal = VacationStartDate.Date.Add(VacationStartTime);
        var endLocal = VacationEndDate.Date.Add(VacationEndTime);
        if (endLocal <= startLocal)
        {
            await _dialog.AlertAsync(
                "Invalid Vacation Window",
                "Vacation end must be after vacation start.");
            return;
        }
        if (endLocal <= DateTime.Now)
        {
            await _dialog.AlertAsync(
                "Invalid Vacation Window",
                "Vacation end is already in the past. Pick a future end date/time.");
            return;
        }

        IsBusy = true;
        try
        {
            // Save state needed to restore the user's normal reply when
            // vacation ends. Persisted (rather than in-memory) so a restart
            // mid-vacation doesn't lose the original text.
            using (_prefs.BeginBatch())
            {
                _prefs.Set("Vacation.Active", true);
                _prefs.Set("Vacation.Start", startLocal.ToString("o"));
                _prefs.Set("Vacation.End", endLocal.ToString("o"));
                // Don't clobber a previously-saved restore snapshot if the
                // user is somehow re-entering vacation while still on it
                // (e.g. extending). Original reply only gets snapshotted on
                // the first transition into vacation.
                if (!IsOnLongVacation)
                {
                    _prefs.Set("Vacation.RestoreInternal", InternalReply);
                    _prefs.Set("Vacation.RestoreExternal", ExternalReply);
                }
            }

            var startOffset = new DateTimeOffset(startLocal, TimeZoneInfo.Local.GetUtcOffset(startLocal));
            var endOffset = new DateTimeOffset(endLocal, TimeZoneInfo.Local.GetUtcOffset(endLocal));

            var settings = new OofSettings
            {
                Status = OofStatus.Scheduled,
                InternalReply = InternalReply,
                ExternalReply = ExternalReply,
                StartTime = startOffset,
                EndTime = endOffset,
            };
            await _exchangeService.SetOofSettingsAsync(settings);
            RememberConfirmedOofState(settings);

            IsOnLongVacation = true;
            // Make sure the checkbox visual matches the just-pushed state
            // even if start was triggered programmatically (e.g. the user
            // pressed ⚡ Sync now with the box already checked and a
            // pre-existing pref).
            IsVacationWindowActive = true;
            _suppressOofToggleCommit = true;
            try
            {
                IsScheduled = true;
                // Long Vacation is a manual OOF intent even when the
                // Scheduled window starts in the future, so the Manual card's
                // switch should read On while the vacation is selected. The
                // top status bar still derives from confirmed Outlook times.
                IsOofEnabled = true;
            }
            finally
            {
                _suppressOofToggleCommit = false;
            }
            CurrentStatus = OofStatus.Scheduled;

            VacationStatus = $"🏖️ On vacation until {endLocal:ddd MMM d, HH:mm}";
            StatusMessage = $"🏖️ Vacation OOF scheduled {startLocal:ddd MMM d HH:mm} → {endLocal:ddd MMM d HH:mm}. Work schedule auto-sync paused.";

            // Keep the automation loop alive even if Work Schedule is off, so
            // the loop can auto-clear vacation when the end time arrives.
            StartAutomationLoop();
        }
        catch (Exception ex)
        {
            // Roll back the prefs flag so a retry isn't blocked by stale state.
            _prefs.Set("Vacation.Active", false);
            IsOnLongVacation = false;
            StatusMessage = $"Failed to start vacation: {ex.Message}";
            await _dialog.AlertAsync("Start Vacation Failed", ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Ends a long-vacation block: clears the multi-day Scheduled OOF window,
    /// restores the user's pre-vacation reply text, and resumes the regular
    /// Work Schedule (if it was on). Called both by the user via the "End
    /// vacation now" button and automatically by the automation loop when the
    /// vacation end time arrives.
    ///
    /// We deliberately make a *single* Exchange round-trip here. Older drafts
    /// did three (Disabled → Enabled/Disabled → Scheduled), which made Outlook
    /// flicker through several states in quick succession and — more
    /// importantly — left a window where prefs said "no vacation" but Exchange
    /// still had the multi-day Scheduled block if the second call failed. The
    /// new flow computes the final post-vacation state up-front, pushes it,
    /// and only clears vacation prefs once Exchange acknowledges the change.
    /// </summary>
    [RelayCommand]
    private async Task EndVacationAsync()
    {
        if (!IsOnLongVacation) return;
        if (IsBusy) return;

        var restoreInternal = _prefs.GetString("Vacation.RestoreInternal", string.Empty) ?? string.Empty;
        var restoreExternal = _prefs.GetString("Vacation.RestoreExternal", string.Empty) ?? string.Empty;

        // Pre-compute the post-vacation target so we make exactly one Set call.
        // ComputeNextOffHoursWindow may return null on a degenerate schedule
        // (every day off-work etc.); in that case we collapse to plain Disabled
        // — same fallback SyncToOutlookCoreAsync already uses.
        OofSettings target;
        if (IsWorkScheduleEnabled)
        {
            var window = ComputeNextOffHoursWindow(DateTime.Now);
            if (window != null)
            {
                target = new OofSettings
                {
                    Status = OofStatus.Scheduled,
                    InternalReply = restoreInternal,
                    ExternalReply = restoreExternal,
                    StartTime = window.Value.start,
                    EndTime = window.Value.end,
                };
            }
            else
            {
                target = new OofSettings
                {
                    Status = OofStatus.Disabled,
                    InternalReply = restoreInternal,
                    ExternalReply = restoreExternal,
                };
            }
        }
        else
        {
            target = new OofSettings
            {
                Status = OofStatus.Disabled,
                InternalReply = restoreInternal,
                ExternalReply = restoreExternal,
            };
        }

        IsBusy = true;
        try
        {
            await _exchangeService.SetOofSettingsAsync(target);
            RememberConfirmedOofState(target);
        }
        catch (Exception ex)
        {
            // Critical: do NOT clear vacation prefs / flip the flag here. Doing
            // so would leave Exchange holding the multi-day Scheduled window
            // while OofManager believed vacation was over — and HasVacationEnded
            // wouldn't be able to recover (it reads from the prefs we'd just
            // wiped). Surface the error and bail; the user can retry.
            StatusMessage = $"End vacation failed: {ex.Message}";
            IsBusy = false;
            return;
        }
        IsBusy = false;

        // Exchange accepted the change. Now it's safe to forget vacation.
        using (_prefs.BeginBatch())
        {
            _prefs.Set("Vacation.Active", false);
            _prefs.Set("Vacation.Start", null);
            _prefs.Set("Vacation.End", null);
        }
        IsOnLongVacation = false;
        // Clear the checkbox too — covers the automation-driven end-of-
        // vacation path where the user never touched the box themselves.
        IsVacationWindowActive = false;

        InternalReply = restoreInternal;
        ExternalReply = restoreExternal;

        // Mirror the pushed state into the local UI flags. Order matches the
        // post-Set bookkeeping in SyncToOutlookCoreAsync / ApplyWorkScheduleAsync
        // so dedupe caches stay in step with what the server now has.
        _suppressOofToggleCommit = true;
        try
        {
            if (target.Status == OofStatus.Scheduled)
            {
                IsScheduled = true;
                StartDate = target.StartTime!.Value.LocalDateTime.Date;
                StartTime = target.StartTime.Value.LocalDateTime.TimeOfDay;
                EndDate = target.EndTime!.Value.LocalDateTime.Date;
                EndTime = target.EndTime.Value.LocalDateTime.TimeOfDay;
                var nowOffset = DateTimeOffset.Now;
                IsOofEnabled = target.StartTime.Value <= nowOffset && target.EndTime.Value > nowOffset;
                CurrentStatus = OofStatus.Scheduled;
            }
            else
            {
                IsScheduled = false;
                IsOofEnabled = false;
                CurrentStatus = OofStatus.Disabled;
            }
        }
        finally
        {
            _suppressOofToggleCommit = false;
        }

        VacationStatus = string.Empty;
        StatusMessage = target.Status == OofStatus.Scheduled && target.StartTime.HasValue && target.EndTime.HasValue
            ? $"🛬 Vacation ended — schedule resumed ({target.StartTime.Value.LocalDateTime:ddd MM-dd HH:mm} → {target.EndTime.Value.LocalDateTime:ddd MM-dd HH:mm})."
            : "🛬 Vacation ended — OOF cleared.";

        if (_tray.IsWindowHidden)
        {
            _tray.ShowNotification(
                "OOF Manager",
                "Vacation ended — out-of-office cleared, regular work schedule resumed.");
        }

        // If Work Schedule is off and we only kept the loop alive for vacation
        // end-detection, shut it back down so we don't poll for nothing —
        // unless Manual mode + Auto-sync still need it for periodic reconcile.
        if (!IsWorkScheduleEnabled && !IsAutoRefreshEnabled)
        {
            StopAutomationLoop();
        }
    }

    /// <summary>
    /// Shared implementation behind both the user's "Sync now" button and the
    /// background auto-sync. Auto-sync calls with isUserInitiated=false so
    /// failures don't pop dialogs and a no-change push doesn't churn the UI.
    /// </summary>
    private async Task SyncToOutlookCoreAsync(bool isUserInitiated)
    {
        if (IsBusy) return;
        if (!IsWorkScheduleEnabled) return;
        // Same reason as ApplyWorkScheduleAsync: don't overwrite the long
        // vacation window with a "next off-hours" stretch.
        if (IsOnLongVacation) return;

        var window = ComputeNextOffHoursWindow(DateTime.Now);
        if (window == null)
        {
            if (isUserInitiated)
            {
                await _dialog.AlertAsync(
                    "Cannot Sync",
                    "There is no upcoming off-hours window in your schedule. Add at least one work day with hours that don't span the full 24 hours.");
            }
            SyncLogger.Write($"SyncCore initiated={(isUserInitiated ? "user" : "auto")} -> no off-hours window in schedule; bail");
            return;
        }
        var (start, end) = window.Value;

        var settings = new OofSettings
        {
            Status = OofStatus.Scheduled,
            InternalReply = InternalReply,
            ExternalReply = ExternalReply,
            StartTime = start,
            EndTime = end
        };
        SyncLogger.Write(
            $"SyncCore initiated={(isUserInitiated ? "user" : "auto")} " +
            $"desired=Scheduled window={start:yyyy-MM-ddTHH:mm}..{end:yyyy-MM-ddTHH:mm}");

        // Server-state verification: before pushing, fetch the live mailbox
        // OOF config and skip the Set only when it already matches what we
        // want. The previous design dedupe'd against a local "last pushed"
        // snapshot, which silently masked external drift — if Outlook
        // desktop's Automatic Replies dialog "OK" button, OWA, an admin, or
        // Outlook Mobile changed the mailbox to Disabled, the 5-min auto-
        // sync would keep no-op'ing because the *intended* payload still
        // matched the cache. Reading the actual server state every tick
        // gives us automatic self-heal at the cost of a single ~50ms Get
        // per check.
        //
        // User-initiated pushes skip this fast-path and always re-assert,
        // so "Check and sync now" is a guaranteed hard re-sync.
        if (!isUserInitiated)
        {
            try
            {
                var current = await _exchangeService.GetOofSettingsAsync();
                // Log a body *fingerprint* (not the plaintext): if a third
                // party (Outlook desktop, Outlook Mobile, OWA, a stray Flow)
                // is the one reverting OOF state, we still want to see in the
                // log whether the body they left behind matches ours — but
                // OOF replies can contain travel plans, phone numbers, or
                // family details, so we never want plaintext on disk. A short
                // SHA-256 prefix is enough to compare "is this the same body
                // we last pushed" without leaking content.
                SyncLogger.Write(
                    $"  pre-Set GET -> {current.Status} " +
                    $"{current.StartTime:yyyy-MM-ddTHH:mm}..{current.EndTime:yyyy-MM-ddTHH:mm} " +
                    $"audience={(current.ExternalAudienceAll ? "All" : "Known")} " +
                    $"intLen={(current.InternalReply ?? string.Empty).Length} " +
                    $"extLen={(current.ExternalReply ?? string.Empty).Length} " +
                    $"intFp={Fingerprint(current.InternalReply)} " +
                    $"extFp={Fingerprint(current.ExternalReply)}");
                if (ServerMatchesDesired(current, settings))
                {
                    SyncLogger.Write("  server already matches desired; no-op");
                    RememberConfirmedOofState(current);
                    return;
                }
            }
            catch (Exception ex)
            {
                SyncLogger.Write($"  pre-Set GET failed: {ex.Message}; falling through to Set");
                // Couldn't verify (network blip, runspace recycled, etc.).
                // Fall through and push — SetOofSettingsAsync does its own
                // read-back verification, so we won't silently succeed on
                // a failure.
            }
        }

        IsBusy = true;
        try
        {
            // SetOofSettingsAsync may now sit through up to a few seconds of
            // write + read-back verification (replica-lag detection). Surface
            // an intermediate caption so the status bar doesn't keep showing
            // a stale "Loading OOF settings..." while we're actually pushing.
            StatusMessage = isUserInitiated
                ? "📤 Syncing to Outlook..."
                : "🔄 Auto-syncing OOF state...";
            await _exchangeService.SetOofSettingsAsync(settings);
            SyncLogger.Write("  Set + read-back verification succeeded");
            RememberConfirmedOofState(settings);

            MarkOofClean(OofStatus.Scheduled);

            // Reflect the new server state locally so the UI's "Out of Office"
            // card matches Outlook (which is now in Scheduled mode). The
            // toggle reflects whether OOF is *actually firing right now* — a
            // future-only Scheduled window means it isn't, so the toggle stays
            // Off until the start time arrives. Without this guard the toggle
            // would visually flip On the moment the user clicks Sync now mid-
            // workday, which lies about whether senders are actually getting
            // auto-replies.
            _suppressOofToggleCommit = true;
            try
            {
                IsScheduled = true;
                var nowOffset = DateTimeOffset.Now;
                IsOofEnabled = start <= nowOffset && end > nowOffset;
            }
            finally
            {
                _suppressOofToggleCommit = false;
            }
            // Re-assert CurrentStatus *after* the IsOofEnabled / IsScheduled
            // setters' partial methods run, so the status label always reads
            // the real mailbox state instead of being clobbered to Disabled
            // when IsOofEnabled is false.
            CurrentStatus = OofStatus.Scheduled;
            StartDate = start.LocalDateTime.Date;
            StartTime = start.LocalDateTime.TimeOfDay;
            EndDate = end.LocalDateTime.Date;
            EndTime = end.LocalDateTime.TimeOfDay;

            var startLabel = start.LocalDateTime.ToString("ddd MM-dd HH:mm");
            var endLabel = end.LocalDateTime.ToString("ddd MM-dd HH:mm");
            var msg = isUserInitiated
                ? $"📤 Outlook will auto-OOF from {startLabel} to {endLabel}. Works without this app open."
                : $"🔄 Auto-sync: Outlook OOF window updated to {startLabel} → {endLabel}";
            StatusMessage = msg;

            // Surface to the tray when the window is hidden so the user still
            // sees the auto-action. The user-initiated push doesn't need a
            // balloon (they're looking at the app and the status bar already
            // updated). Also dedupe against the last *notified* window so the
            // 5-min self-heal loop doesn't spam balloons every tick — the user
            // only sees one when the next off-hours window actually moves.
            if (!isUserInitiated
                && _tray.IsWindowHidden
                && (_lastNotifiedStart != start || _lastNotifiedEnd != end))
            {
                _tray.ShowNotification(
                    "OOF Manager",
                    $"Outlook auto-reply scheduled: {startLabel} → {endLabel}");
                _lastNotifiedStart = start;
                _lastNotifiedEnd = end;
            }
        }
        catch (Exception ex)
        {
            SyncLogger.Write($"  Set FAILED: {ex.GetType().Name}: {ex.Message}");
            // Surface manual failures loudly; auto failures only update the
            // status bar so we don't pop a dialog from a background tick.
            StatusMessage = $"Sync to Outlook failed: {ex.Message}";
            if (isUserInitiated)
            {
                await _dialog.AlertAsync("Sync Failed", ex.Message);
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Last-chance push invoked from <c>App.OnExit</c>: re-asserts the current
    /// Scheduled off-hours window so the mailbox stays in the right state for
    /// the period after OofManager closes. Anything that flipped the mailbox
    /// to Disabled while we were running (Outlook desktop's Automatic Replies
    /// dialog, OWA, an admin) gets corrected here. Best-effort and bounded by
    /// the caller's wait timeout — failures are swallowed because we don't
    /// have a UI thread to surface them on during shutdown.
    ///
    /// Gated on <see cref="IsAutoRefreshEnabled"/>: when auto-refresh is off,
    /// the user has explicitly opted out of OofManager mutating Outlook in the
    /// background, so we must not race a Scheduled write into Exchange after
    /// the window closes (the user might be tweaking Automatic Replies in
    /// Outlook right after closing us). Toggle ON, startup, and exit are all
    /// no-push by default; mailbox writes only happen on explicit
    /// ⚡ Sync now / 💾 Save changes, or when the user enables Auto-refresh.
    /// </summary>
    public async Task FlushBeforeExitAsync()
    {
        // On vacation, the multi-day Scheduled window already covers shutdown;
        // never reassert anything that could overwrite it.
        if (IsOnLongVacation) return;
        // Auto-sync gates background mutation in both modes. Off → no
        // last-chance push on exit (the user has explicitly opted out of
        // OofManager writing to Exchange in the background).
        if (!IsAutoRefreshEnabled) return;
        if (!_exchangeService.IsConnected) return;

        if (IsManualMode)
        {
            // Manual mode: re-assert current toggle state + reply text. The
            // helper short-circuits when the server already matches, so this
            // is cheap when nothing drifted.
            try { await ReassertManualStateAsync(isUserInitiated: false); } catch { }
            return;
        }

        var window = ComputeNextOffHoursWindow(DateTime.Now);
        if (window == null) return;
        var (start, end) = window.Value;

        var settings = new OofSettings
        {
            Status = OofStatus.Scheduled,
            InternalReply = InternalReply,
            ExternalReply = ExternalReply,
            StartTime = start,
            EndTime = end,
        };

        try
        {
            await _exchangeService.SetOofSettingsAsync(settings);
        }
        catch
        {
            // Shutdown path — nowhere to surface this. SetOofSettingsAsync
            // already verifies via read-back, so a swallowed exception here
            // genuinely means the server didn't accept the change; the next
            // launch's LoadAsync + sync will re-establish the correct state.
        }
    }

    /// <summary>
    /// Computes the next contiguous off-hours window starting at or after now,
    /// based on the per-day work-hour configuration. Returns null if no such
    /// window exists (e.g. every day is off-work, leaving no anchor for end).
    /// </summary>
    private (DateTimeOffset start, DateTimeOffset end)? ComputeNextOffHoursWindow(DateTime now)
    {
        // Off-hours start: anchor it to the most recent end-of-work boundary,
        // not to "now". Without this, opening the app on Saturday morning
        // (or pre-9am on Mon) would push a window starting at the moment of
        // launch — but the user has been off-work since Friday 17:30, so the
        // logical start is Friday 17:30. Walking backwards keeps the window
        // honest about when OOF actually began.
        DateTime offStart;
        if (IsWorkday(now.DayOfWeek))
        {
            var endToday = now.Date.Add(GetEndTimeForDay(now.DayOfWeek));
            var startToday = now.Date.Add(GetStartTimeForDay(now.DayOfWeek));
            if (now < startToday)
            {
                // Pre-work-hours on a workday (e.g. 07:30 on a Mon with 09:00
                // start). The off-hours stretch began at the previous workday's
                // end-of-work; only fall back to "now" if we genuinely can't
                // find any earlier workday in the schedule.
                offStart = FindMostRecentEndOfWorkAtOrBefore(now) ?? now;
            }
            else
            {
                // Inside or after today's work hours: anchor to today's end.
                // (FindNextStartOfWorkAfter walks past today, so end == next
                // workday's start, which is what we want.)
                offStart = endToday;
            }
        }
        else
        {
            // Today is off-work entirely (Sat/Sun by default). Anchor to the
            // last workday's end-of-work — that's when off-hours actually
            // started — falling back to "now" only if no workday is configured.
            offStart = FindMostRecentEndOfWorkAtOrBefore(now) ?? now;
        }

        // Off-hours end: the very next start-of-work boundary strictly after
        // offStart. Walk forward up to 7 days; bail if no day is a workday.
        DateTime? offEnd = FindNextStartOfWorkAfter(offStart);
        if (offEnd == null || offEnd.Value <= offStart) return null;

        return (
            new DateTimeOffset(offStart, TimeZoneInfo.Local.GetUtcOffset(offStart)),
            new DateTimeOffset(offEnd.Value, TimeZoneInfo.Local.GetUtcOffset(offEnd.Value))
        );
    }

    private DateTime? FindNextStartOfWorkAfter(DateTime t)
    {
        // Today first: only counts if we haven't passed today's start yet.
        if (IsWorkday(t.DayOfWeek))
        {
            var startToday = t.Date.Add(GetStartTimeForDay(t.DayOfWeek));
            if (t < startToday) return startToday;
        }
        for (int i = 1; i <= 7; i++)
        {
            var nextDate = t.Date.AddDays(i);
            if (IsWorkday(nextDate.DayOfWeek))
                return nextDate.Add(GetStartTimeForDay(nextDate.DayOfWeek));
        }
        return null;
    }

    private DateTime? FindMostRecentEndOfWorkAtOrBefore(DateTime t)
    {
        // Today first: only counts if we've already crossed today's end.
        if (IsWorkday(t.DayOfWeek))
        {
            var endToday = t.Date.Add(GetEndTimeForDay(t.DayOfWeek));
            if (t >= endToday) return endToday;
        }
        for (int i = 1; i <= 7; i++)
        {
            var prevDate = t.Date.AddDays(-i);
            if (IsWorkday(prevDate.DayOfWeek))
                return prevDate.Add(GetEndTimeForDay(prevDate.DayOfWeek));
        }
        return null;
    }

    /// <summary>
    /// Re-evaluates "are we inside working hours right now?" and refreshes
    /// the schedule card's caption. The actual mailbox push lives in
    /// <see cref="SyncToOutlookCoreAsync"/>, which every caller of this
    /// method invokes immediately afterwards — pushing flat Enabled/Disabled
    /// here first would briefly transition Exchange through an Enabled-
    /// with-expired-window or Disabled-with-expired-window state, both of
    /// which Outlook desktop's OOF banner reads as "automatic replies are
    /// off" and would flicker the banner for ~1–2s on every tick. Leaving
    /// the mailbox in continuous Scheduled state is much cleaner. The
    /// caption is still updated locally so the schedule card stays in sync
    /// with "now" without waiting for the Outlook push to complete.
    /// </summary>
    private Task ApplyWorkScheduleAsync(bool showSuccessMessage, bool suppressTrayNotification = false)
    {
        // Parameters are kept for call-site compatibility; the function no
        // longer pushes a flat state, so they're informational only. The
        // status-bar / tray balloons happen inside SyncToOutlookCoreAsync
        // (which is always invoked after this) when an actual change ships.
        _ = showSuccessMessage;
        _ = suppressTrayNotification;

        if (!IsWorkScheduleEnabled || IsBusy) return Task.CompletedTask;
        // Long-vacation mode owns the OOF window outright; the regular work
        // schedule must not touch state under it.
        if (IsOnLongVacation) return Task.CompletedTask;

        WorkScheduleStatus = GetWorkScheduleStatus();
        RefreshOofStatusBar();
        return Task.CompletedTask;
    }

    [RelayCommand]
    private async Task SaveOofAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        StatusMessage = "Saving reply messages...";

        try
        {
            var settings = new OofSettings
            {
                Status = _confirmedOofStatus,
                InternalReply = InternalReply,
                ExternalReply = ExternalReply,
                StartTime = _confirmedOofStartTime,
                EndTime = _confirmedOofEndTime,
            };

            await _exchangeService.SetOofSettingsAsync(settings);
            MarkOofClean(settings.Status);
            RememberConfirmedOofState(settings);
            StatusMessage = "✅ Reply messages updated in Outlook.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Save failed: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task SaveAsTemplateAsync()
    {
        var name = await _dialog.PromptAsync(
            "Save Template", "Enter a name for the template:", "Save", "Cancel", placeholder: "e.g. Vacation, Travel, Meeting");

        if (string.IsNullOrWhiteSpace(name)) return;
        var trimmedName = name!.Trim();

        var template = new OofTemplate
        {
            Name = trimmedName,
            InternalReply = InternalReply,
            ExternalReply = ExternalReply,
        };

        await _templateService.SaveTemplateAsync(template);
        // Insert at top (UpdatedAt order) without re-querying the DB. Insert/Remove
        // on ObservableCollection triggers an incremental UI update.
        Templates.Insert(0, template);
        StatusMessage = $"✅ Template \u201c{trimmedName}\u201d saved";
    }

    [RelayCommand]
    private async Task DeleteTemplateAsync(OofTemplate template)
    {
        if (template == null) return;
        var confirm = await _dialog.ConfirmAsync(
            "Delete Template", $"Are you sure you want to delete the template \u201c{template.Name}\u201d?", "Delete", "Cancel");

        if (!confirm) return;

        await _templateService.DeleteTemplateAsync(template.Id);
        Templates.Remove(template);
        StatusMessage = $"Template \u201c{template.Name}\u201d deleted";
    }

    [RelayCommand]
    private async Task SwitchAccountAsync()
    {
        StopAutomationLoop();
        // Reset the load gate so the next sign-in refreshes data instead of
        // showing stale state.
        _hasLoadedOnce = false;
        // Forget the cached UPN so neither the silent token-cache hit nor the
        // Windows-account fallback will silently re-sign-in the same user. The
        // intent here is explicitly "let me pick a different account".
        _prefs.Set("Auth.LastSignedInUpn", null);
        await _exchangeService.DisconnectAsync();
        // forceAccountPicker=true tells the navigation layer to immediately
        // kick off WAM with no UPN hint, so the account picker comes up
        // instead of dropping the user back to the Sign In button.
        _navigation.NavigateToLogin(forceAccountPicker: true);
    }

    private void RefreshOofStatusBar()
    {
        if (!_hasLoadedOnce && IsBusy)
        {
            OofStatusBarMessage = "Loading OOF status...";
            return;
        }

        OofStatusBarMessage = BuildOofStatusBarMessage();
    }

    private string BuildOofStatusBarMessage()
    {
        var now = DateTimeOffset.Now;

        if (IsConfirmedOofActive(now))
        {
            var end = GetCurrentOofEndTime(now);
            return end.HasValue
                ? $"✅ OOF is ON until {FormatStatusDateTime(end.Value)}"
                : "✅ OOF is ON — no end time";
        }

        var nextWindow = GetNextOofWindowForStatusBar(now);
        return nextWindow.HasValue
            ? $"☑️ OOF is OFF — next on {FormatStatusDateTime(nextWindow.Value.start)} to {FormatStatusDateTime(nextWindow.Value.end)}"
            : "☑️ OOF is OFF — no upcoming automatic OOF";
    }

    private DateTimeOffset? GetCurrentOofEndTime(DateTimeOffset now)
    {
        return _confirmedOofStatus == OofStatus.Scheduled
            && _confirmedOofEndTime.HasValue
            && _confirmedOofEndTime.Value > now
                ? _confirmedOofEndTime
                : null;
    }

    private (DateTimeOffset start, DateTimeOffset end)? GetNextOofWindowForStatusBar(DateTimeOffset now)
    {
        // The top status bar is confirmed Outlook state only. Do not preview
        // the local weekly schedule here; otherwise simply switching from
        // Manual mode to Schedule mode changes the status line before the
        // user clicks Sync now or Auto-sync actually writes anything.
        if (_confirmedOofStatus == OofStatus.Scheduled
            && _confirmedOofStartTime.HasValue
            && _confirmedOofEndTime.HasValue
            && _confirmedOofEndTime.Value > now)
        {
            return (_confirmedOofStartTime.Value, _confirmedOofEndTime.Value);
        }

        return null;
    }

    private bool IsConfirmedOofActive(DateTimeOffset now)
    {
        if (_confirmedOofStatus == OofStatus.Enabled) return true;
        if (_confirmedOofStatus != OofStatus.Scheduled) return false;
        return _confirmedOofStartTime.HasValue
            && _confirmedOofStartTime.Value <= now
            && (!_confirmedOofEndTime.HasValue || _confirmedOofEndTime.Value > now);
    }

    private void RememberConfirmedOofState(OofSettings settings)
        => RememberConfirmedOofState(settings.Status, settings.StartTime, settings.EndTime);

    private void RememberConfirmedOofState(OofStatus status, DateTimeOffset? start = null, DateTimeOffset? end = null)
    {
        _confirmedOofStatus = status;
        _confirmedOofStartTime = start;
        _confirmedOofEndTime = end;
        RefreshOofStatusBar();
    }

    private static string FormatStatusDateTime(DateTimeOffset value)
        => value.LocalDateTime.ToString("ddd MM-dd HH:mm");

    private bool IsNowInsideWorkingHours(DateTime now)
    {
        if (!IsWorkday(now.DayOfWeek)) return false;
        var start = GetStartTimeForDay(now.DayOfWeek);
        var end = GetEndTimeForDay(now.DayOfWeek);
        return now.TimeOfDay >= start && now.TimeOfDay < end;
    }

    private bool IsWorkday(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => IsMondayWorkday,
        DayOfWeek.Tuesday => IsTuesdayWorkday,
        DayOfWeek.Wednesday => IsWednesdayWorkday,
        DayOfWeek.Thursday => IsThursdayWorkday,
        DayOfWeek.Friday => IsFridayWorkday,
        DayOfWeek.Saturday => IsSaturdayWorkday,
        DayOfWeek.Sunday => IsSundayWorkday,
        _ => false
    };

    private TimeSpan GetStartTimeForDay(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => MondayStartTime,
        DayOfWeek.Tuesday => TuesdayStartTime,
        DayOfWeek.Wednesday => WednesdayStartTime,
        DayOfWeek.Thursday => ThursdayStartTime,
        DayOfWeek.Friday => FridayStartTime,
        DayOfWeek.Saturday => SaturdayStartTime,
        DayOfWeek.Sunday => SundayStartTime,
        _ => TimeSpan.Zero
    };

    private TimeSpan GetEndTimeForDay(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => MondayEndTime,
        DayOfWeek.Tuesday => TuesdayEndTime,
        DayOfWeek.Wednesday => WednesdayEndTime,
        DayOfWeek.Thursday => ThursdayEndTime,
        DayOfWeek.Friday => FridayEndTime,
        DayOfWeek.Saturday => SaturdayEndTime,
        DayOfWeek.Sunday => SundayEndTime,
        _ => TimeSpan.Zero
    };

    private static string GetDayDisplayName(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => "Monday",
        DayOfWeek.Tuesday => "Tuesday",
        DayOfWeek.Wednesday => "Wednesday",
        DayOfWeek.Thursday => "Thursday",
        DayOfWeek.Friday => "Friday",
        DayOfWeek.Saturday => "Saturday",
        DayOfWeek.Sunday => "Sunday",
        _ => day.ToString()
    };

    private static readonly DayOfWeek[] WeekDays =
    {
        DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
        DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday
    };

    private string GetWorkScheduleStatus()
    {
        if (!IsWorkScheduleEnabled) return "Work schedule rule disabled";
        if (IsOnLongVacation) return "🏖️ Paused — on long vacation";
        return IsNowInsideWorkingHours(DateTime.Now)
            ? "Currently inside working hours: OOF should be off"
            : "Currently outside working hours: OOF should be on";
    }

    private void LoadWorkSchedulePreferences()
    {
        _suppressDirtyTracking = true;
        try
        {
            // Always start in Schedule mode. Users can switch to Manual mode
            // for the current session, but the next app launch should return
            // to the schedule-first workflow.
            IsWorkScheduleEnabled = true;
            IsAutoRefreshEnabled = _prefs.GetBool("WorkSchedule.AutoRefresh", false);
            IsMondayWorkday = _prefs.GetBool("WorkSchedule.Monday", true);
            IsTuesdayWorkday = _prefs.GetBool("WorkSchedule.Tuesday", true);
            IsWednesdayWorkday = _prefs.GetBool("WorkSchedule.Wednesday", true);
            IsThursdayWorkday = _prefs.GetBool("WorkSchedule.Thursday", true);
            IsFridayWorkday = _prefs.GetBool("WorkSchedule.Friday", true);
            IsSaturdayWorkday = _prefs.GetBool("WorkSchedule.Saturday", true);
            IsSundayWorkday = _prefs.GetBool("WorkSchedule.Sunday", true);

            // Backward compatibility: older builds stored a single global window in
            // "WorkSchedule.StartMinutes" / "WorkSchedule.EndMinutes". Use it as the
            // default for every per-day key that hasn't been written yet, so existing
            // users see their previous schedule replicated across all days on first
            // launch of the per-day-aware build.
            var legacyStart = _prefs.GetInt("WorkSchedule.StartMinutes", 9 * 60);
            var legacyEnd = _prefs.GetInt("WorkSchedule.EndMinutes", 18 * 60);

            MondayStartTime = LoadDayTime("Monday", "Start", legacyStart);
            MondayEndTime = LoadDayTime("Monday", "End", legacyEnd);
            TuesdayStartTime = LoadDayTime("Tuesday", "Start", legacyStart);
            TuesdayEndTime = LoadDayTime("Tuesday", "End", legacyEnd);
            WednesdayStartTime = LoadDayTime("Wednesday", "Start", legacyStart);
            WednesdayEndTime = LoadDayTime("Wednesday", "End", legacyEnd);
            ThursdayStartTime = LoadDayTime("Thursday", "Start", legacyStart);
            ThursdayEndTime = LoadDayTime("Thursday", "End", legacyEnd);
            FridayStartTime = LoadDayTime("Friday", "Start", legacyStart);
            FridayEndTime = LoadDayTime("Friday", "End", legacyEnd);
            SaturdayStartTime = LoadDayTime("Saturday", "Start", legacyStart);
            SaturdayEndTime = LoadDayTime("Saturday", "End", legacyEnd);
            SundayStartTime = LoadDayTime("Sunday", "Start", legacyStart);
            SundayEndTime = LoadDayTime("Sunday", "End", legacyEnd);

            // Restore vacation state. We don't push anything to Exchange here —
            // LoadAsync will see IsOnLongVacation=true and skip the regular
            // schedule apply/sync. Auto-end (if the end time has already passed)
            // is handled by the automation loop's HasVacationEnded check.
            IsOnLongVacation = _prefs.GetBool("Vacation.Active", false);
            // The checkbox state persists separately so a planned-but-not-yet-
            // synced vacation survives a restart. Falls back to IsOnLongVacation
            // for users whose prefs predate this key.
            IsVacationWindowActive = _prefs.GetBool("Vacation.WindowActive", IsOnLongVacation);
            var vacationStartRaw = _prefs.GetString("Vacation.Start");
            var vacationEndRaw = _prefs.GetString("Vacation.End");
            if (DateTime.TryParse(vacationStartRaw,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeLocal | System.Globalization.DateTimeStyles.AdjustToUniversal,
                out var startUtc))
            {
                var startLocal = startUtc.ToLocalTime();
                VacationStartDate = startLocal.Date;
                VacationStartTime = startLocal.TimeOfDay;
            }
            if (DateTime.TryParse(vacationEndRaw,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeLocal | System.Globalization.DateTimeStyles.AdjustToUniversal,
                out var endUtc))
            {
                var endLocal = endUtc.ToLocalTime();
                VacationEndDate = endLocal.Date;
                VacationEndTime = endLocal.TimeOfDay;
                if (IsOnLongVacation)
                {
                    VacationStatus = $"🏖️ On vacation until {endLocal:ddd MMM d, HH:mm}";
                }
            }

            WorkScheduleStatus = GetWorkScheduleStatus();
        }
        finally
        {
            _suppressDirtyTracking = false;
            HasUnsavedScheduleChanges = false;
        }
    }

    private TimeSpan LoadDayTime(string day, string suffix, int legacyDefaultMinutes)
    {
        var minutes = _prefs.GetInt($"WorkSchedule.{day}.{suffix}Minutes", legacyDefaultMinutes);
        // Snap the persisted values to the same 30-minute grid the dropdown
        // exposes; otherwise an off-grid value (from an older build that
        // accepted free text) would make the ComboBox show blank.
        return SnapToHalfHour(TimeSpan.FromMinutes(minutes));
    }

    private static TimeSpan SnapToHalfHour(TimeSpan value)
    {
        var totalMinutes = value.TotalMinutes;
        var snapped = (int)Math.Round(totalMinutes / 30.0) * 30;
        if (snapped < 0) snapped = 0;
        if (snapped > 23 * 60 + 30) snapped = 23 * 60 + 30;
        return TimeSpan.FromMinutes(snapped);
    }

    private void SaveWorkSchedulePreferences()
    {
        // Batch all writes into a single disk flush.
        using (_prefs.BeginBatch())
        {
            _prefs.Set("WorkSchedule.Enabled", IsWorkScheduleEnabled);
            _prefs.Set("WorkSchedule.AutoRefresh", IsAutoRefreshEnabled);
            _prefs.Set("WorkSchedule.Monday", IsMondayWorkday);
            _prefs.Set("WorkSchedule.Tuesday", IsTuesdayWorkday);
            _prefs.Set("WorkSchedule.Wednesday", IsWednesdayWorkday);
            _prefs.Set("WorkSchedule.Thursday", IsThursdayWorkday);
            _prefs.Set("WorkSchedule.Friday", IsFridayWorkday);
            _prefs.Set("WorkSchedule.Saturday", IsSaturdayWorkday);
            _prefs.Set("WorkSchedule.Sunday", IsSundayWorkday);

            _prefs.Set("WorkSchedule.Monday.StartMinutes", (int)MondayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Monday.EndMinutes", (int)MondayEndTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Tuesday.StartMinutes", (int)TuesdayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Tuesday.EndMinutes", (int)TuesdayEndTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Wednesday.StartMinutes", (int)WednesdayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Wednesday.EndMinutes", (int)WednesdayEndTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Thursday.StartMinutes", (int)ThursdayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Thursday.EndMinutes", (int)ThursdayEndTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Friday.StartMinutes", (int)FridayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Friday.EndMinutes", (int)FridayEndTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Saturday.StartMinutes", (int)SaturdayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Saturday.EndMinutes", (int)SaturdayEndTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Sunday.StartMinutes", (int)SundayStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.Sunday.EndMinutes", (int)SundayEndTime.TotalMinutes);
        }
        WorkScheduleStatus = GetWorkScheduleStatus();
    }

    private void StartAutomationLoop()
    {
        if (_automationCts != null) return;
        _automationCts = new CancellationTokenSource();
        _ = RunAutomationLoopAsync(_automationCts.Token);
    }

    private void StopAutomationLoop()
    {
        _automationCts?.Cancel();
        _automationCts?.Dispose();
        _automationCts = null;
    }

    private async Task RunAutomationLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                // Sleep until either the next work-hours boundary or 5 minutes
                // (whichever is sooner). The boundary-aligned wake guarantees
                // we flip OOF *exactly* at start/end times, while the 5-minute
                // fallback self-heals against clock changes, workday config
                // changes, and out-of-band edits to the mailbox.
                var delay = ComputeNextCheckDelay(DateTime.Now);
                await Task.Delay(delay, cancellationToken);
                // Marshal to the UI thread and *await the inner Task to actual completion*.
                // Dispatcher.InvokeAsync(Func<Task>) returns DispatcherOperation<Task>; awaiting
                // that only waits for the lambda to hit its first yield, not for the inner work
                // to finish. Unwrapping forces a full wait so two iterations can never overlap.
                var op = Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    // Time-based vacation end runs unconditionally — losing
                    // this tick would leave Exchange holding stale OOF after
                    // the user is back. (User-intent end, i.e. unchecking the
                    // box, is folded into the Manual reassert path below so
                    // it respects the auto-sync gate.)
                    if (IsOnLongVacation && HasVacationEnded(DateTime.Now))
                    {
                        // Background-sync intent: when the vacation expires
                        // automatically, the user wants the weekly schedule to
                        // pick up where it left off — not Manual mode + plain
                        // Disabled. If we're still in Manual mode (the typical
                        // path: user kicked off Vacation/Holiday from there),
                        // flip the underlying schedule flag *first* so the
                        // EndVacationAsync branch that pushes the next off-
                        // hours Scheduled window runs. Suppress the commit
                        // handler so we don't pop a validation alert on a
                        // grid the user never touched; ComputeNextOffHoursWindow
                        // already falls back to plain Disabled on a degenerate
                        // schedule, so this is safe even when the grid is
                        // empty.
                        if (!IsWorkScheduleEnabled)
                        {
                            _suppressWorkScheduleCommit = true;
                            try { IsWorkScheduleEnabled = true; }
                            finally { _suppressWorkScheduleCommit = false; }
                            SaveWorkSchedulePreferences();
                        }
                        await EndVacationAsync();
                        return;
                    }
                    // Auto-sync off → the periodic loop is for vacation-end
                    // detection only. We don't touch the mailbox until the user
                    // takes another explicit action (toggle / save changes /
                    // ⚡ Sync now).
                    if (!IsAutoRefreshEnabled) return;

                    if (IsScheduleMode)
                    {
                        // Vacation block owns OOF; never overwrite it with a
                        // weekly-schedule push.
                        if (IsOnLongVacation) return;
                        await ApplyWorkScheduleAsync(showSuccessMessage: false, suppressTrayNotification: true);
                        // Auto-sync runs after the local toggle so Outlook stays
                        // in lock-step with the client's view, and so the
                        // *next* off-hours window is already pre-pushed before
                        // the user's working day even ends.
                        await SyncToOutlookCoreAsync(isUserInitiated: false);
                    }
                    else
                    {
                        // Manual mode + auto-sync: reconcile current local
                        // state against what the server holds, correcting
                        // drift from other clients (Outlook desktop, OWA,
                        // an admin, Power Automate flows). Also pushes
                        // pending vacation start/end transitions when the
                        // checkbox diverges from IsOnLongVacation.
                        await ReassertManualStateAsync(isUserInitiated: false);
                    }
                });
                await op.Task.Unwrap();
            }
        }
        catch (OperationCanceledException) { }
    }

    /// <summary>
    /// True when the persisted vacation end timestamp is at or before
    /// <paramref name="now"/>. Reads from prefs (rather than the in-memory
    /// VacationEndDate/Time pair) so that an end time written by a previous
    /// process / a previous launch still drives auto-end correctly.
    /// </summary>
    private bool HasVacationEnded(DateTime now)
    {
        var raw = _prefs.GetString("Vacation.End");
        if (string.IsNullOrEmpty(raw)) return true;
        if (!DateTime.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AssumeLocal | System.Globalization.DateTimeStyles.AdjustToUniversal,
            out var endUtc))
        {
            return true;
        }
        // Convert to local for comparison with DateTime.Now (which is local).
        var endLocal = endUtc.ToLocalTime();
        return now >= endLocal;
    }

    private TimeSpan ComputeNextCheckDelay(DateTime now)
    {
        // Cap the wait to 5 minutes so we self-heal if the clock changes,
        // workdays change, etc.
        var fallback = TimeSpan.FromMinutes(5);
        TimeSpan? candidate = null;

        // Today's boundaries (only relevant if today is itself a workday).
        if (IsWorkday(now.DayOfWeek))
        {
            var startToday = now.Date.Add(GetStartTimeForDay(now.DayOfWeek));
            var endToday = now.Date.Add(GetEndTimeForDay(now.DayOfWeek));
            if (now < startToday) candidate = startToday - now;
            else if (now < endToday) candidate = endToday - now;
        }

        // Past today's window (or today is not a workday) — find the next
        // enabled workday and aim at its start time. Capped at 7 days to
        // avoid an infinite loop if the user disables every day.
        if (candidate == null)
        {
            for (int i = 1; i <= 7; i++)
            {
                var nextDate = now.Date.AddDays(i);
                if (IsWorkday(nextDate.DayOfWeek))
                {
                    candidate = nextDate.Add(GetStartTimeForDay(nextDate.DayOfWeek)) - now;
                    break;
                }
            }
        }

        var delay = candidate.HasValue && candidate.Value < fallback ? candidate.Value : fallback;
        // Don't poll faster than once a minute regardless.
        return delay < TimeSpan.FromMinutes(1) ? TimeSpan.FromMinutes(1) : delay;
    }

    /// <summary>
    /// True when the server's current OOF config already matches the payload
    /// we'd push. Used by <see cref="SyncToOutlookCoreAsync"/> to detect
    /// (and self-heal) external drift while skipping the Set round-trip on
    /// steady-state ticks.
    /// </summary>
    private static bool ServerMatchesDesired(OofSettings server, OofSettings desired)
    {
        if (server.Status != desired.Status) return false;
        if (server.ExternalAudienceAll != desired.ExternalAudienceAll) return false;

        // The Set path writes unzoned local datetime strings that Exchange
        // interprets in the mailbox's timezone; reading them back goes
        // through a different parser. The combination can introduce sub-
        // minute jitter even when neither side moved, so compare at minute
        // precision in UTC to avoid a spurious push every tick.
        if (!TimestampsMatch(server.StartTime, desired.StartTime)) return false;
        if (!TimestampsMatch(server.EndTime, desired.EndTime)) return false;

        // Reply text goes through PlainTextToHtml on Set and HtmlToPlainText
        // on Get; the roundtrip can normalize line endings and trim trailing
        // whitespace. Compare a normalized form so that doesn't churn us
        // into pushing every tick.
        if (!RepliesMatch(server.InternalReply, desired.InternalReply)) return false;
        if (!RepliesMatch(server.ExternalReply, desired.ExternalReply)) return false;

        return true;
    }

    private static bool TimestampsMatch(DateTimeOffset? a, DateTimeOffset? b)
    {
        if (!a.HasValue && !b.HasValue) return true;
        if (!a.HasValue || !b.HasValue) return false;
        return Math.Abs((a.Value.UtcDateTime - b.Value.UtcDateTime).TotalMinutes) < 1.0;
    }

    private static bool RepliesMatch(string? a, string? b)
    {
        static string Normalize(string? s) => (s ?? string.Empty)
            .Replace("\r\n", "\n").Trim();
        return string.Equals(Normalize(a), Normalize(b), StringComparison.Ordinal);
    }

    /// <summary>
    /// Short hex fingerprint of an OOF reply body for sync.log. Same input
    /// always produces the same value (so we can still answer "is the body
    /// drifting between ticks?") but the original plaintext never lands on
    /// disk \u2014 OOF replies may contain travel plans, phone numbers, or
    /// family details that don't belong in a diagnostic log.
    /// </summary>
    private static string Fingerprint(string? value)
    {
        if (string.IsNullOrEmpty(value)) return "<empty>";
        using var sha = System.Security.Cryptography.SHA256.Create();
        var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(value!));
        // 4 bytes / 8 hex chars: collisions are theoretically possible but
        // we only need to distinguish "same as last tick" vs "something
        // changed", and the body length we log alongside helps disambiguate.
        var sb = new System.Text.StringBuilder(8);
        for (int i = 0; i < 4; i++) sb.Append(bytes[i].ToString("x2"));
        return sb.ToString();
    }
}
