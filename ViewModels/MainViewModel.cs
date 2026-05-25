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
    private CancellationTokenSource? _automationCts;
    private bool _hasLoadedOnce;
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
    // Manual-mode toggle can now diverge locally until Sync now
    // pushes it, so the top line must not be derived from IsOofEnabled.
    private OofStatus _confirmedOofStatus = OofStatus.Disabled;
    private DateTimeOffset? _confirmedOofStartTime;
    private DateTimeOffset? _confirmedOofEndTime;
    private bool _isCloudScheduleFlowStatusChecking;
    private bool _isCloudSchedulePackageRunning;
    private bool _isVacationFlowsStatusChecking;
    // Set to true around any programmatic mutation of IsOofEnabled so the
    // partial setter's auto-commit path (which only fires for genuine user
    // gestures on the OOF toggle) doesn't kick a Set against Exchange when
    // we're just reflecting server state locally (LoadAsync, vacation
    // start/end, explicit sync pushes, etc.).
    private bool _suppressOofToggleCommit;

    [ObservableProperty] private bool _isBusy;
    // Operation/status feedback for button-driven work (Sync now, package
    // generation, save failures, etc.). This is shown inside the OOF Settings
    // card, not in the persistent top status bar.
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private string _cloudScheduleFlowBannerText = "Power Automate flow: Not checked. Use the buttons to check or change the cloud schedule flow.";
    [ObservableProperty] private string _cloudScheduleFlowStateText = "Not checked";
    [ObservableProperty] private string _cloudScheduleFlowBannerDetail = "Use the buttons to check or change the cloud schedule flow.";
    // Manual-vacation cloud flows banner — sibling to the Cloud Schedule
    // banner but reports the combined state of the two flows (Vacation Start
    // + Vacation End). Combined state values: "On" (both On), "Off" (both
    // Off), "Mixed" (one On one Off), "Not found" (neither imported yet),
    // "Partial" (only one of the two found), "Checking", "Unknown".
    [ObservableProperty] private string _vacationFlowsStateText = "Not checked";
    [ObservableProperty] private string _vacationFlowsBannerDetail = "Use the buttons to check or change the cloud vacation flows.";
    // Persistent top status bar: always describes the real/current OOF state
    // plus the next known OOF window, never transient button progress.
    [ObservableProperty] private string _oofStatusBarMessage = "Loading OOF status...";
    [ObservableProperty] private OofStatus _currentStatus = OofStatus.Disabled;
    [ObservableProperty] private bool _isOofEnabled;
    [ObservableProperty] private bool _isScheduled;
    [ObservableProperty] private bool _isWorkScheduleEnabled;
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
    /// save. The OOF switch is intentionally deferred to Sync now,
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
    /// ⚡ Sync now button will actually do in the currently-selected mode, so
    /// they don't have to guess whether "sync" means "push the schedule" or
    /// "push my manual state".
    /// </summary>
    public string SyncCardSubtitle
    {
        get
        {
            if (IsScheduleMode)
                return "Weekly mode: OOF Manager auto-flips OOF based on the weekly hours below. ⚡ Sync to Outlook pushes the next off-hours window immediately.";
            return "Manual mode: you flip OOF on/off yourself below. ⚡ Sync to Outlook pushes your current state to Outlook immediately.";
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
    // vacation cleanup and explicit sync stop fighting Exchange and the OOF window
    // is the single multi-day Scheduled block we pushed in StartVacationAsync.
    // Pre-vacation reply text is squirreled away in prefs so EndVacation can
    // restore the user's normal reply without forcing them to re-type it.
    [ObservableProperty] private bool _isOnLongVacation;
    // User-intent counterpart to IsOnLongVacation. The checkbox in the Manual
    // mode card toggles this; the actual push to Exchange happens later, when
    // the user clicks ⚡ Sync now
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
        IPreferencesService prefs)
    {
        _exchangeService = exchangeService;
        _powerAutomate = powerAutomate;
        _templateService = templateService;
        _dialog = dialog;
        _navigation = navigation;
        _prefs = prefs;
    }

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
        // only. Exchange is updated by ⚡ Sync now
        // path (ReassertManualStateAsync). Programmatic sets wrap their
        // writes in _suppressOofToggleCommit so they don't show a pending
        // action note.
        if (_suppressOofToggleCommit) return;
        if (IsOofAutoManaged || !_hasLoadedOnce || IsBusy) return;
        if (IsOnLongVacation)
        {
            StatusMessage = value
                ? "Vacation OOF window remains selected in Outlook. Click ⚡ Sync to Outlook to re-assert it."
                : "Vacation OOF window cleared locally. Click ⚡ Sync to Outlook to clear it in Outlook.";
            return;
        }
        StatusMessage = value
            ? "OOF switch set to ON locally. Click ⚡ Sync to Outlook to push it to Outlook."
            : "OOF switch set to OFF locally. Click ⚡ Sync to Outlook to push it to Outlook.";
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
    // deferred until the user clicks ⚡ Sync now
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
                await _exchangeService.TryAutoConnectAsync(lastUpn!, TimeSpan.FromSeconds(45));
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
            var displayName = await _exchangeService.GetCurrentDisplayNameAsync();
            MailboxIdentity = mailboxIdentity;
            UserDisplayName = string.IsNullOrWhiteSpace(displayName) || displayName.Contains("@")
                ? mailboxIdentity
                : displayName;
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
            // the next off-hours window pre-push would flip the
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
                // We deliberately do NOT push to Outlook on launch; mailbox
                // writes now happen only from explicit user actions or the
                // Power Automate cloud schedule.
                IsBusy = false;
                await ApplyWorkScheduleAsync(showSuccessMessage: false);
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
            // Refresh Cloud Schedule flow status on every load so the banner
            // (now shown in Schedule mode) reflects the live M365 state at
            // app start. Previously gated on IsManualMode because the banner
            // lived in Manual mode as a "your cloud flow is still running"
            // reminder; now that it's a live status control in Schedule
            // mode it should reflect reality whenever the app is up.
            if (_hasLoadedOnce)
            {
                _ = RefreshCloudScheduleFlowStatusAsync();
                _ = RefreshVacationFlowsStatusAsync();
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
            // Mode changes are local UI/model changes. The OOF Settings subtitle
            // already explains what Sync to Outlook will do, so don't echo
            // another status line underneath it.
            if (!_isCloudSchedulePackageRunning) StatusMessage = string.Empty;
            await Task.CompletedTask;
        }
        else
        {
            if (!IsOnLongVacation)
            {
                StopAutomationLoop();
            }
            WorkScheduleStatus = "Work schedule rule disabled";

            // Switching to Manual mode is now a local mode change only. Do
            // not collapse an existing Scheduled window to flat Enabled /
            // Disabled here; that was the path that made the top status jump
            // to "OOF is ON — no end time" before the user clicked Sync now.
            if (!_isCloudSchedulePackageRunning) StatusMessage = string.Empty;
        }
    }

    [RelayCommand]
    private async Task SaveWorkScheduleAsync()
    {
        if (!await TrySavePendingScheduleChangesAsync()) return;

        if (IsWorkScheduleEnabled)
        {
            await ApplyWorkScheduleAsync(showSuccessMessage: true);
            await SyncToOutlookCoreAsync(isUserInitiated: true);
        }
        else
        {
            if (!IsOnLongVacation)
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
                StatusMessage = "Clearing vacation OOF window and syncing the schedule...";
                await EndVacationAsync();
                if (IsOnLongVacation) return;
                StatusMessage = "Weekly mode synced. Vacation OOF window cleared.";
                return;
            }

            if (IsVacationWindowActive)
            {
                IsVacationWindowActive = false;
                VacationStatus = string.Empty;
            }

            // 1. Local re-check + Exchange OOF flip.
            await ApplyWorkScheduleAsync(showSuccessMessage: false);

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
    /// button and non-user sync paths; the latter sets
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
                : "🔄 Syncing OOF state...";
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
    /// Builds the OofManager Cloud Schedule solution package under local app
    /// data, then tries the automatic Power Automate import path. If the
    /// import tool is unavailable, sign-in fails, or the import fails, it still opens
    /// make.powerautomate.com/solutions so the user can drop the zip into the
    /// Import dialog manually.
    /// </summary>
    [RelayCommand]
    private async Task GenerateCloudSchedulePackageAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        _isCloudSchedulePackageRunning = true;
        try
        {
            var snapshot = new WorkScheduleSnapshot(_prefs);
            var packageDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OofManager",
                "CloudSchedule");
            var outPath = System.IO.Path.Combine(packageDir, "OofManager-CloudSchedule.zip");

            StatusMessage = "📦 Generating cloud schedule solution…";
            var pkg = await Task.Run(() => CloudSchedulePackageGenerator.GenerateWithIdentity(
                snapshot,
                userEmail: MailboxIdentity,
                internalReply: InternalReply,
                externalReply: ExternalReply,
                externalAudienceAll: true,
                generateManaged: false,
                outputPath: outPath));

            StatusMessage = $"📦 Generated cloud schedule solution v{pkg.SolutionVersion}. Importing to Power Automate…";
            var importProgress = new Progress<string>(message =>
            {
                if (!string.IsNullOrWhiteSpace(message))
                    StatusMessage = $"📦 {message}";
            });
            var import = await _powerAutomate.ImportCloudScheduleSolutionAsync(
                pkg.Path,
                pkg.SolutionUniqueName,
                pkg.WorkflowId,
                pkg.FlowDisplayName,
                MailboxIdentity,
                UserDisplayName,
                forceOverwrite: false,
                progress: importProgress);

            if (import.Outcome == CloudScheduleImportOutcome.Success)
            {
                var activation = await EnsureCloudScheduleFlowOnAfterImportAsync(pkg.FlowDisplayName);
                var prefix = activation.IsReady ? "✅" : "⚠️";
                StatusMessage = $"{prefix} {import.Message} Version {pkg.SolutionVersion}; imported at {DateTime.Now:yyyy-MM-dd HH:mm}. {activation.Message}";
                return;
            }

            OpenSolutionsPage(import.EnvironmentId);
            StatusMessage = $"⚠️ Automatic Power Automate import did not finish: {import.Message} Power Automate opened — import the saved zip manually from {pkg.Path}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to generate cloud schedule solution: {ex.Message}";
            await _dialog.AlertAsync("Cloud Schedule Solution", ex.Message);
        }
        finally
        {
            _isCloudSchedulePackageRunning = false;
            IsBusy = false;
        }
    }

    /// <summary>
    /// Manual-vacation cloud orchestrator (research/manual-vacation-cloud-flows
    /// branch). Mirrors <see cref="GenerateCloudSchedulePackageAsync"/>: builds
    /// the Manual-Vacation solution zip, runs the same Power Automate import
    /// path against the user's owned env, then turns ON both generated flows
    /// (Vacation Start + Vacation End) so they actually fire at their one-shot
    /// trigger times. Falls back to opening the maker portal + showing the
    /// saved zip path if any step fails, same recovery shape the Schedule
    /// import uses.
    /// </summary>
    [RelayCommand]
    private async Task GenerateManualVacationPackageAsync()
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
                "Vacation Already Past",
                "The vacation end is in the past — the one-shot triggers would never fire.");
            return;
        }

        IsBusy = true;
        try
        {
            // Pull the existing Schedule flow's runtime ids from the same prefs
            // PowerAutomateService writes after the user's first successful
            // toggle. See repo memory note on FlowName != WorkflowId. Without
            // these, the zip ships with AutoReply only (no pause/resume of the
            // Schedule flow).
            var scheduleEnvId = _prefs.GetString("PowerAutomate.Flow.Environment.Id");
            var scheduleFlowName = _prefs.GetString("PowerAutomate.Flow.Name");
            var hasScheduleTarget = !string.IsNullOrWhiteSpace(scheduleEnvId) && !string.IsNullOrWhiteSpace(scheduleFlowName);

            var packageDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OofManager",
                "ManualVacation");
            var outPath = System.IO.Path.Combine(packageDir, "OofManager-ManualVacation.zip");

            StatusMessage = "🏖️ Generating manual vacation solution…";
            var pkg = await Task.Run(() => ManualVacationPackageGenerator.GenerateWithIdentity(
                userEmail: MailboxIdentity,
                vacationStart: startLocal,
                vacationEnd: endLocal,
                internalReply: InternalReply,
                externalReply: ExternalReply,
                externalAudienceAll: true,
                scheduleFlowEnvironmentId: scheduleEnvId,
                scheduleFlowRuntimeFlowName: scheduleFlowName,
                generateManaged: false,
                outputPath: outPath));

            StatusMessage = $"🏖️ Generated manual vacation v{pkg.SolutionVersion}. Importing to Power Automate…";
            var importProgress = new Progress<string>(message =>
            {
                if (!string.IsNullOrWhiteSpace(message))
                    StatusMessage = $"🏖️ {message}";
            });

            // Reuses the generic Cloud Schedule import pipeline — env discovery,
            // pac auth, ImportSolutionAsync, PublishAllXml — by passing the
            // manual-vacation solution name + the start flow's deterministic
            // WorkflowId (used only as a hint; the script echoes it back).
            var import = await _powerAutomate.ImportCloudScheduleSolutionAsync(
                pkg.Path,
                pkg.SolutionUniqueName,
                pkg.StartWorkflowId,
                pkg.StartFlowDisplayName,
                MailboxIdentity,
                UserDisplayName,
                forceOverwrite: false,
                progress: importProgress);

            if (import.Outcome != CloudScheduleImportOutcome.Success)
            {
                OpenSolutionsPage(import.EnvironmentId);
                StatusMessage = $"⚠️ Automatic manual vacation import did not finish: {import.Message} Power Automate opened — import the saved zip manually from {pkg.Path}.";
                return;
            }

            var startOn = await EnsureManualVacationFlowOnAsync(pkg.StartFlowDisplayName, "Vacation Start");
            var endOn   = await EnsureManualVacationFlowOnAsync(pkg.EndFlowDisplayName, "Vacation End");

            var bothOn = startOn.IsReady && endOn.IsReady;
            if (!bothOn)
            {
                // Most common cause: the shared_flowmanagement connection
                // reference is unbound because the user has never created a
                // Power Automate Management connection on this tenant.
                // pac CLI / Dataverse import can't bind it automatically.
                // Open the solution's Connection references page directly so
                // the user is one click away from the fix.
                OpenConnectionReferencesPage(import.EnvironmentId, pkg.SolutionUniqueName);
                StatusMessage = $"⚠️ Manual vacation v{pkg.SolutionVersion} imported but flows are still Off. " +
                                $"Start: {startOn.Message} End: {endOn.Message}";
                await _dialog.AlertAsync("Bind connections to turn the vacation flows on",
                    "Both Vacation flows were imported but couldn't be turned on. Usually this means a connection reference inside the solution isn't bound to a connection yet — most often 'OofManager Flow Management' (the Power Automate Management connector), because it's the first time you're using it on this tenant.\n\n" +
                    "I just opened the Connection references page. For each row whose Status is Off:\n" +
                    "  1. Click the row's ⋯ → Edit\n" +
                    "  2. In Connection: pick an existing one, or click '+ New connection' and sign in\n" +
                    "  3. Save\n\n" +
                    "Then come back to the Vacation flows (Solutions → OofManager Manual Vacation → Cloud flows) and click 'Turn on'. After this one-time setup, re-running 'Set up vacation in cloud' will reuse the bound connections automatically.");
                return;
            }

            var prefix = "✅";
            var pauseNote = hasScheduleTarget
                ? "Pause/resume of your Cloud Schedule flow is wired."
                : "AutoReply-only build — toggle the Cloud Schedule flow once first if you also want pause/resume.";
            StatusMessage = $"{prefix} Manual vacation v{pkg.SolutionVersion} imported. Start flow: {startOn.Message} End flow: {endOn.Message} {pauseNote}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to set up manual vacation: {ex.Message}";
            await _dialog.AlertAsync("Manual Vacation", ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Turns on a freshly-imported manual-vacation flow. Calls Enable
    /// directly instead of status-then-enable: a flow that's already On
    /// makes Enable a cheap no-op, and a freshly-imported flow can briefly
    /// be invisible to <c>Get-Flow</c> while Power Automate finishes
    /// registering it — the status check returning <c>NotFound</c> in that
    /// window would otherwise make us bail without ever calling Enable.
    /// One quick retry after a short delay covers the registration lag.
    /// </summary>
    private async Task<(bool IsReady, string Message)> EnsureManualVacationFlowOnAsync(string expectedFlowDisplayName, string label)
    {
        var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
        var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
        var statusProgress = new Progress<string>(message =>
        {
            if (!string.IsNullOrWhiteSpace(message))
                StatusMessage = $"🏖️ {label}: {message}";
        });

        async Task<PowerAutomateResult> TryEnableOnceAsync() =>
            await _powerAutomate.EnableOofManagerFlowsAsync(upn, displayName, expectedFlowDisplayName, progress: statusProgress);

        try
        {
            var enable = await TryEnableOnceAsync();
            if (enable.Outcome == PowerAutomateOutcome.NoFlowFound)
            {
                // Freshly-imported flow may not be queryable for a few
                // seconds. Wait + retry once.
                StatusMessage = $"🏖️ {label}: flow not visible yet, waiting 10s before retry…";
                await Task.Delay(TimeSpan.FromSeconds(10));
                enable = await TryEnableOnceAsync();
            }
            return enable.Outcome == PowerAutomateOutcome.Success
                ? (true, $"{label} turned on.")
                : (false, $"{label} could not be turned on: {enable.Message}");
        }
        catch (Exception ex)
        {
            return (false, $"{label} enable failed: {ex.Message}");
        }
    }

    private async Task<(bool IsReady, string Message)> EnsureCloudScheduleFlowOnAfterImportAsync(string expectedFlowDisplayName)
    {
        var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
        var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;

        SetCloudScheduleFlowBannerProgress("Checking", "Checking Power Automate flow after import...");
        StatusMessage = "📦 Import completed. Checking Power Automate flow status...";

        var statusProgress = new Progress<string>(message =>
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            var detail = NormalizeCloudScheduleFlowDetail(message);
            SetCloudScheduleFlowBannerProgress("Checking", detail);
            StatusMessage = $"📦 {detail}";
        });

        try
        {
            var status = await _powerAutomate.GetOofManagerFlowStatusAsync(upn, displayName, expectedFlowDisplayName, progress: statusProgress);
            SetCloudScheduleFlowBanner(status.State);

            if (status.State == PowerAutomateFlowState.On)
            {
                return (true, "Power Automate flow is on.");
            }

            if (status.State != PowerAutomateFlowState.Off)
            {
                return (false, "Power Automate solution imported, but the flow status could not be confirmed: " + status.Message);
            }

            SetCloudScheduleFlowBannerProgress("Turning on", "Power Automate flow is off. Turning it on...");
            StatusMessage = "📦 Power Automate flow is off. Turning it on...";

            var enableProgress = new Progress<string>(message =>
            {
                if (string.IsNullOrWhiteSpace(message)) return;
                StatusMessage = message;
            });
            var enable = await _powerAutomate.EnableOofManagerFlowsAsync(upn, displayName, expectedFlowDisplayName, enableProgress);

            switch (enable.Outcome)
            {
                case PowerAutomateOutcome.Success:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.On);
                    return (true, "Power Automate flow was off and has been turned on.");
                case PowerAutomateOutcome.NoFlowFound:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.NotFound);
                    OpenPowerAutomateFlows();
                    return (false, enable.Message);
                case PowerAutomateOutcome.SignInFailed:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
                    OpenPowerAutomateFlows();
                    return (false, "Power Automate sign-in did not complete, so the flow could not be turned on automatically.");
                case PowerAutomateOutcome.SolutionAwareBlocked:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
                    OpenPowerAutomateFlows();
                    return (false, "Power Automate refused the automatic turn-on. Opened Power Automate so you can turn it on there.");
                default:
                    SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
                    OpenPowerAutomateFlows();
                    return (false, enable.Message);
            }
        }
        catch (Exception ex)
        {
            SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
            return (false, "Power Automate solution imported, but the flow status check failed: " + ex.Message);
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
    /// Deep-links to the Connection references view of a specific solution
    /// in the maker portal. Used after a manual-vacation import when one of
    /// the connection references is unbound (shared_flowmanagement is the
    /// usual culprit on fresh tenants) — drops the user one click away from
    /// the binding UI instead of forcing them to navigate
    /// Solutions → solution → Objects → Connection references by hand.
    /// </summary>
    private static void OpenConnectionReferencesPage(string? environmentId, string solutionUniqueName)
    {
        if (string.IsNullOrWhiteSpace(environmentId) || string.IsNullOrWhiteSpace(solutionUniqueName))
        {
            OpenSolutionsPage(environmentId);
            return;
        }
        // The maker portal honors a query-string filter on the connections
        // landing page (?solution={uniqueName}); from there the connection-
        // refs panel is one click. There isn't a stable direct URL into the
        // connection-references list of a solution that survives portal
        // refactors, so this is the most reliable shape.
        var url = $"https://make.powerautomate.com/environments/{environmentId}/solutions/{Uri.EscapeDataString(solutionUniqueName)}";
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

    /// <summary>
    /// Compares the locally-configured Weekly schedule (workdays + earliest
    /// end-of-day time used as the cloud Recurrence trigger) against what's
    /// actually deployed in the user's Cloud Schedule flow. Reports the diff
    /// in a dialog so the user can spot drift between local and cloud (e.g.
    /// after editing on a different machine, or after fiddling in the Power
    /// Automate designer). Per-day work hours aren't compared in v1 — they're
    /// encoded inside Logic Apps nested-if expressions and need either a
    /// reverse parser or a sidecar metadata field on the generator side.
    /// </summary>
    [RelayCommand]
    private async Task CompareCloudScheduleAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        StatusMessage = "☁️ Comparing local schedule against cloud flow…";
        try
        {
            var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
            var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
            var expectedFlowDisplayName = CloudSchedulePackageGenerator.ComputeFlowIdentity(upn ?? string.Empty).FlowDisplayName;
            var progress = new Progress<string>(m => { if (!string.IsNullOrWhiteSpace(m)) StatusMessage = $"☁️ {m}"; });

            var result = await _powerAutomate.GetCloudScheduleDefinitionAsync(upn, displayName, expectedFlowDisplayName, progress);

            if (result.Outcome == PowerAutomateOutcome.NoFlowFound)
            {
                StatusMessage = "☁️ No cloud flow found — import one first.";
                await _dialog.AlertAsync("Compare with cloud",
                    "No OofManager Cloud Schedule flow was found in your Power Platform environment. Click 'Generate & Import Solution' first.");
                return;
            }
            if (result.Outcome != PowerAutomateOutcome.Success)
            {
                StatusMessage = $"☁️ Compare failed: {result.Message}";
                await _dialog.AlertAsync("Compare with cloud", $"Couldn't read the cloud flow definition.\n\n{result.Message}");
                return;
            }

            var weekDays = new[] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday };
            var localWorkdays = weekDays.Where(IsWorkday).ToList();
            TimeSpan? localTriggerEnd = null;
            foreach (var d in localWorkdays)
            {
                var e = GetEndTimeForDay(d);
                if (localTriggerEnd == null || e < localTriggerEnd.Value) localTriggerEnd = e;
            }

            string CloudTriggerLocalString()
            {
                if (result.TriggerHour is not int h || result.TriggerMinute is not int m) return "(not set)";
                try
                {
                    var local = CloudTriggerToLocal(h, m, result.TriggerTimeZone);
                    return $"{local:hh\\:mm} (local) — cloud raw: {h:D2}:{m:D2} {result.TriggerTimeZone ?? "UTC"}";
                }
                catch
                {
                    return $"{h:D2}:{m:D2} {result.TriggerTimeZone ?? "UTC"}";
                }
            }

            string FmtDays(IEnumerable<DayOfWeek> days) =>
                string.Join(", ", weekDays.Where(days.Contains).Select(d => d.ToString().Substring(0, 3)));

            var localDaysStr = localWorkdays.Count == 0 ? "(none)" : FmtDays(localWorkdays);
            var cloudDaysStr = result.WorkDays.Count == 0 ? "(none)" : FmtDays(result.WorkDays);
            var localTriggerStr = localTriggerEnd.HasValue ? localTriggerEnd.Value.ToString(@"hh\:mm") + " (local)" : "(no workdays)";

            var workdaysMatch = localWorkdays.OrderBy(d => d).SequenceEqual(result.WorkDays.OrderBy(d => d));

            bool triggerMatches = false;
            if (result.TriggerHour is int ch && result.TriggerMinute is int cm && localTriggerEnd is TimeSpan let)
            {
                try
                {
                    var cloudLocal = CloudTriggerToLocal(ch, cm, result.TriggerTimeZone);
                    triggerMatches = cloudLocal.Hours == let.Hours && cloudLocal.Minutes == let.Minutes;
                }
                catch { /* leave false */ }
            }

            var allMatch = workdaysMatch && triggerMatches;
            var headline = allMatch ? "✅ Local and cloud match." : "⚠️ Local and cloud differ.";

            var body = headline + "\n\n" +
                       $"Flow: {result.FlowDisplayName ?? expectedFlowDisplayName}\n\n" +
                       $"  Workdays:\n" +
                       $"    Local: {localDaysStr}\n" +
                       $"    Cloud: {cloudDaysStr}   {(workdaysMatch ? "✓" : "✗")}\n\n" +
                       $"  Earliest workday end (= cloud trigger time):\n" +
                       $"    Local: {localTriggerStr}\n" +
                       $"    Cloud: {CloudTriggerLocalString()}   {(triggerMatches ? "✓" : "✗")}\n\n" +
                       "Per-day hours aren't compared in v1 (encoded inside Logic Apps expressions).\n\n" +
                       (allMatch
                          ? "Everything we can check looks consistent."
                          : "If local is the authoritative version, click 'Generate & Import Solution' to push it to the cloud.");

            StatusMessage = allMatch ? "☁️ Local and cloud match." : "☁️ Local and cloud differ — see dialog.";
            await _dialog.AlertAsync("Compare with cloud", body);
        }
        catch (Exception ex)
        {
            StatusMessage = $"☁️ Compare failed: {ex.Message}";
            await _dialog.AlertAsync("Compare with cloud", ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Converts the cloud Recurrence trigger's raw hour/minute (which is
    /// expressed in <paramref name="triggerTimeZone"/>, not UTC) into local
    /// time-of-day for comparison against the local schedule. The generator
    /// stamps the trigger timeZone as the user's local <see cref="TimeZoneInfo.Local"/>
    /// id at import time, so most of the time triggerTimeZone == local and
    /// the conversion is a no-op; the explicit conversion still matters for
    /// the cross-machine / cross-tz case (user imports from one PC and
    /// later checks from another in a different zone).
    /// </summary>
    private static TimeSpan CloudTriggerToLocal(int hour, int minute, string? triggerTimeZone)
    {
        var srcTz = TimeZoneInfo.Local;
        if (!string.IsNullOrWhiteSpace(triggerTimeZone))
        {
            try { srcTz = TimeZoneInfo.FindSystemTimeZoneById(triggerTimeZone!); }
            catch
            {
                if (string.Equals(triggerTimeZone, "UTC", StringComparison.OrdinalIgnoreCase))
                    srcTz = TimeZoneInfo.Utc;
            }
        }
        if (srcTz.Id == TimeZoneInfo.Local.Id)
            return new TimeSpan(hour, minute, 0);
        var unspec = new DateTime(DateTime.UtcNow.Year, DateTime.UtcNow.Month, DateTime.UtcNow.Day, hour, minute, 0, DateTimeKind.Unspecified);
        var asLocal = TimeZoneInfo.ConvertTime(unspec, srcTz, TimeZoneInfo.Local);
        return asLocal.TimeOfDay;
    }

    private async Task RefreshCloudScheduleFlowStatusAsync()
    {
        if (_isCloudScheduleFlowStatusChecking) return;

        _isCloudScheduleFlowStatusChecking = true;
        SetCloudScheduleFlowBannerProgress("Checking", "Checking Power Automate flow...");
        try
        {
            var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
            var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
            var expectedFlowDisplayName = CloudSchedulePackageGenerator.ComputeFlowIdentity(upn ?? string.Empty).FlowDisplayName;
            var statusProgress = new Progress<string>(message =>
            {
                if (string.IsNullOrWhiteSpace(message)) return;
                SetCloudScheduleFlowBannerProgress("Checking", NormalizeCloudScheduleFlowDetail(message));
            });
            var result = await _powerAutomate.GetOofManagerFlowStatusAsync(upn, displayName, expectedFlowDisplayName, progress: statusProgress);
            SetCloudScheduleFlowBanner(result.State);
        }
        catch
        {
            SetCloudScheduleFlowBanner(PowerAutomateFlowState.Unknown);
        }
        finally
        {
            _isCloudScheduleFlowStatusChecking = false;
        }
    }

    private void SetCloudScheduleFlowBanner(PowerAutomateFlowState state)
    {
        var (stateText, detailText) = state switch
        {
            PowerAutomateFlowState.On => ("On", "Turn it off before vacation; turn it back on after."),
            PowerAutomateFlowState.Off => ("Off", "Turn it back on after vacation to resume the cloud schedule."),
            PowerAutomateFlowState.NotFound => ("Not found", "Import the Cloud Schedule package first."),
            _ => ("Unknown", "Sign in to Power Automate to check or use the buttons."),
        };

        CloudScheduleFlowStateText = stateText;
        CloudScheduleFlowBannerDetail = detailText;
        CloudScheduleFlowBannerText = $"Power Automate flow: {stateText}. {detailText}";
    }

    private void SetCloudScheduleFlowBannerProgress(string stateText, string detailText)
    {
        CloudScheduleFlowStateText = stateText;
        CloudScheduleFlowBannerDetail = detailText;
        CloudScheduleFlowBannerText = $"Power Automate flow: {detailText}";
    }

    private static string NormalizeCloudScheduleFlowDetail(string message)
    {
        const string prefix = "Power Automate flow:";
        var detail = message.Trim();
        return detail.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            ? detail.Substring(prefix.Length).TrimStart()
            : detail;
    }

    /// <summary>
    /// Manual-mode counterpart to <see cref="DisableCloudScheduleFlowAsync"/>:
    /// turns OFF both Vacation Start and Vacation End flows in one click.
    /// Used when the user wants to cancel a planned cloud vacation without
    /// deleting the imported solution (so the flows are dormant but can be
    /// re-enabled later via the Turn-on button).
    /// </summary>
    [RelayCommand]
    private Task DisableVacationFlowsAsync() => RunVacationFlowsToggleAsync(disable: true);

    /// <summary>
    /// Re-enables both Vacation Start and Vacation End flows. Used after a
    /// previous Turn-off, or to recover from a Power Automate import that
    /// left the flows Off because a connection reference was unbound.
    /// </summary>
    [RelayCommand]
    private Task EnableVacationFlowsAsync() => RunVacationFlowsToggleAsync(disable: false);

    private async Task RefreshVacationFlowsStatusAsync()
    {
        if (_isVacationFlowsStatusChecking) return;

        _isVacationFlowsStatusChecking = true;
        SetVacationFlowsBannerProgress("Checking", "Checking vacation flows...");
        try
        {
            var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
            var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
            var identity = ManualVacationPackageGenerator.ComputeIdentity(upn ?? string.Empty);

            var startStatus = await _powerAutomate.GetOofManagerFlowStatusAsync(upn, displayName, identity.StartFlowDisplayName);
            var endStatus   = await _powerAutomate.GetOofManagerFlowStatusAsync(upn, displayName, identity.EndFlowDisplayName);
            SetVacationFlowsBanner(startStatus.State, endStatus.State);
        }
        catch
        {
            SetVacationFlowsBanner(PowerAutomateFlowState.Unknown, PowerAutomateFlowState.Unknown);
        }
        finally
        {
            _isVacationFlowsStatusChecking = false;
        }
    }

    private void SetVacationFlowsBanner(PowerAutomateFlowState startState, PowerAutomateFlowState endState)
    {
        // Combine the two states into one user-facing label. Distinguishing
        // Mixed vs Partial helps the user understand whether the next step
        // is "re-enable one flow" vs "(re-)import the vacation solution".
        string state, detail;
        if (startState == PowerAutomateFlowState.NotFound && endState == PowerAutomateFlowState.NotFound)
        {
            state = "Not found";
            detail = "No cloud vacation planned. Pick dates above and click 🏖️ Set up vacation in cloud.";
        }
        else if (startState == PowerAutomateFlowState.NotFound || endState == PowerAutomateFlowState.NotFound)
        {
            state = "Partial";
            detail = "Only one of the two vacation flows was found. Re-run 🏖️ Set up vacation in cloud to repair.";
        }
        else if (startState == PowerAutomateFlowState.On && endState == PowerAutomateFlowState.On)
        {
            state = "On";
            detail = "Both vacation flows are armed. They'll fire at your start/end times.";
        }
        else if (startState == PowerAutomateFlowState.Off && endState == PowerAutomateFlowState.Off)
        {
            state = "Off";
            detail = "Both vacation flows exist but are off. Click Turn on to arm them.";
        }
        else if (startState == PowerAutomateFlowState.On || endState == PowerAutomateFlowState.On)
        {
            state = "Mixed";
            detail = $"Start: {startState}, End: {endState}. Click Turn on to arm both.";
        }
        else
        {
            state = "Unknown";
            detail = "Sign in to Power Automate to check, or use the buttons.";
        }

        VacationFlowsStateText = state;
        VacationFlowsBannerDetail = detail;
    }

    private void SetVacationFlowsBannerProgress(string stateText, string detailText)
    {
        VacationFlowsStateText = stateText;
        VacationFlowsBannerDetail = detailText;
    }

    private async Task RunVacationFlowsToggleAsync(bool disable)
    {
        if (IsBusy) return;
        var verbLabel = disable ? "Turning off" : "Turning on";
        IsBusy = true;
        StatusMessage = $"{verbLabel} vacation flows...";
        SetVacationFlowsBannerProgress(disable ? "Turning off" : "Turning on", $"{verbLabel} both vacation flows...");
        try
        {
            var upn = !string.IsNullOrWhiteSpace(MailboxIdentity) ? MailboxIdentity : UserEmail;
            var displayName = !string.IsNullOrWhiteSpace(UserDisplayName) ? UserDisplayName : null;
            var identity = ManualVacationPackageGenerator.ComputeIdentity(upn ?? string.Empty);

            async Task<PowerAutomateResult> ToggleOneAsync(string flowDisplayName, string label)
            {
                SetVacationFlowsBannerProgress(disable ? "Turning off" : "Turning on", $"{verbLabel} {label}...");
                var progress = new Progress<string>(m =>
                {
                    if (!string.IsNullOrWhiteSpace(m))
                        StatusMessage = $"{label}: {m}";
                });
                return disable
                    ? await _powerAutomate.DisableOofManagerFlowsAsync(upn, displayName, flowDisplayName, progress: progress)
                    : await _powerAutomate.EnableOofManagerFlowsAsync(upn, displayName, flowDisplayName, progress: progress);
            }

            var startResult = await ToggleOneAsync(identity.StartFlowDisplayName, "Vacation Start");
            var endResult   = await ToggleOneAsync(identity.EndFlowDisplayName, "Vacation End");

            await RefreshVacationFlowsStatusAsync();

            var bothOk = startResult.Outcome == PowerAutomateOutcome.Success && endResult.Outcome == PowerAutomateOutcome.Success;

            // When the user is trying to TURN ON vacation flows and the
            // PowerShell call reports Success but the flows are still Off
            // after a refresh, the near-certain cause is that the
            // shared_flowmanagement connection reference is unbound: pac CLI
            // / Dataverse import can't auto-bind it on tenants where the
            // user has never created a Power Automate Management connection
            // before (which is everyone the first time). Power Automate
            // accepts the Enable API call but refuses to actually activate
            // a flow with an unbound connection ref, so the UI symptom is
            // "Turn on does nothing." Deep-link to the solution's Connection
            // references page so the user is one click from binding it.
            if (!disable && bothOk && VacationFlowsStateText != "On")
            {
                var envId = _prefs.GetString("PowerAutomate.Import.Environment.Id");
                OpenConnectionReferencesPage(envId, identity.SolutionUniqueName);
                StatusMessage = "⚠️ Vacation flows accepted the Turn on but stayed Off — Connection references page opened.";
                await _dialog.AlertAsync("Bind the Flow Management connection",
                    "The vacation flows accepted the Turn on call but stayed Off. The usual cause is an unbound 'OofManager Flow Management' connection reference (Power Automate Management connector). This is a one-time setup per tenant.\n\n" +
                    "I just opened the Connection references page. For each row whose Status is Off:\n" +
                    "  1. Click the row's ⋯ → Edit\n" +
                    "  2. In Connection: click '+ New connection' → sign in\n" +
                    "  3. Save\n\n" +
                    "Then come back here and click '▶️ Turn on vacation flows' again.");
                return;
            }

            StatusMessage = bothOk
                ? $"✅ Vacation flows {(disable ? "turned off" : "turned on")}."
                : $"⚠️ Vacation flows toggle finished with issues. Start: {startResult.Outcome} ({startResult.Message}). End: {endResult.Outcome} ({endResult.Message}).";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Vacation flows toggle failed: {ex.Message}";
            SetVacationFlowsBanner(PowerAutomateFlowState.Unknown, PowerAutomateFlowState.Unknown);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task RunCloudScheduleFlowToggleAsync(bool disable)
    {
        if (IsBusy) return;
        var verbLabel = disable ? "Turning off" : "Turning on";
        IsBusy = true;
        StatusMessage = $"{verbLabel} Power Automate flow. Preparing...";
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
            var toggleProgress = new Progress<string>(message =>
            {
                if (!string.IsNullOrWhiteSpace(message))
                    StatusMessage = message;
            });
            var result = disable
                ? await _powerAutomate.DisableOofManagerFlowsAsync(upn, displayName, expectedFlowDisplayName, toggleProgress)
                : await _powerAutomate.EnableOofManagerFlowsAsync(upn, displayName, expectedFlowDisplayName, toggleProgress);
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
    /// the regular Work Schedule rules so it doesn't fight the
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
            StatusMessage = $"🏖️ Vacation OOF scheduled {startLocal:ddd MMM d HH:mm} → {endLocal:ddd MMM d HH:mm}. Work schedule paused.";

            // Keep a lightweight vacation watcher alive while the app is open,
            // so it can auto-clear vacation when the end time arrives.
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

        StopAutomationLoop();
    }

    /// <summary>
    /// Shared implementation behind both the user's "Sync now" button and the
    /// non-user sync paths. Those calls use isUserInitiated=false so
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
                : "🔄 Syncing OOF state...";
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
                : $"📤 Outlook OOF window updated to {startLabel} → {endLabel}";
            StatusMessage = msg;

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
    private Task ApplyWorkScheduleAsync(bool showSuccessMessage)
    {
        // Parameters are kept for call-site compatibility; the function no
        // longer pushes a flat state, so they're informational only. The
        // status bar is updated inside SyncToOutlookCoreAsync when an actual
        // change ships.
        _ = showSuccessMessage;

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
                ? $"✅ Outlook OOF is ON until {FormatStatusDateTime(end.Value)}"
                : "✅ Outlook OOF is ON — no end time";
        }

        var nextWindow = GetNextOofWindowForStatusBar(now);
        return nextWindow.HasValue
            ? $"☑️ Outlook OOF is OFF — next scheduled {FormatStatusDateTime(nextWindow.Value.start)} to {FormatStatusDateTime(nextWindow.Value.end)}"
            : "☑️ Outlook OOF is OFF — no upcoming Outlook scheduled OOF";
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
        // user clicks Sync now.
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

            // Restore vacation state. We don't push anything to Exchange here.
            // LoadAsync will auto-end only if the stored end time has passed.
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
                var delay = ComputeNextVacationCheckDelay(DateTime.Now);
                await Task.Delay(delay, cancellationToken);
                // Marshal to the UI thread and *await the inner Task to actual completion*.
                // Dispatcher.InvokeAsync(Func<Task>) returns DispatcherOperation<Task>; awaiting
                // that only waits for the lambda to hit its first yield, not for the inner work
                // to finish. Unwrapping forces a full wait so two iterations can never overlap.
                var op = Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    // This loop now exists only for vacation auto-end while
                    // the app is open. Regular work schedule advancement is
                    // handled by explicit Sync or Power Automate.
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

    private TimeSpan ComputeNextVacationCheckDelay(DateTime now)
    {
        var fallback = TimeSpan.FromMinutes(5);
        var raw = _prefs.GetString("Vacation.End");
        if (DateTime.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AssumeLocal | System.Globalization.DateTimeStyles.AdjustToUniversal,
            out var endUtc))
        {
            var untilEnd = endUtc.ToLocalTime() - now;
            if (untilEnd > TimeSpan.Zero && untilEnd < fallback)
                return untilEnd < TimeSpan.FromSeconds(5) ? TimeSpan.FromSeconds(5) : untilEnd;
        }
        return fallback;
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
