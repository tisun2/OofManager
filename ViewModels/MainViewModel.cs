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
    private readonly ITemplateService _templateService;
    private readonly IDialogService _dialog;
    private readonly INavigationService _navigation;
    private readonly IPreferencesService _prefs;
    private readonly ITrayService _tray;
    private readonly IStartupService _startup;
    private CancellationTokenSource? _automationCts;
    private bool _hasLoadedOnce;
    // Last off-hours window pushed to Outlook by SyncToOutlookCoreAsync. Used to
    // skip a redundant server round-trip when the auto-sync loop computes the
    // same window twice in a row (the common case during a long off-hours stretch).
    private DateTimeOffset? _lastSyncedStart;
    private DateTimeOffset? _lastSyncedEnd;
    // Last (start,end) we actually showed a tray balloon for. Tracked separately
    // from the sync cache because ApplyWorkScheduleAsync invalidates the sync
    // cache after every successful SetOofSettingsAsync (which is correct — each
    // call wipes any Scheduled window on Exchange, so we must re-push it). If we
    // gated the balloon on the same cache, the user would get a balloon every
    // 3-min self-heal tick. This pair only updates when the *window itself*
    // actually changes, so the user only sees a balloon for meaningful moves.
    private DateTimeOffset? _lastNotifiedStart;
    private DateTimeOffset? _lastNotifiedEnd;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private OofStatus _currentStatus = OofStatus.Disabled;
    [ObservableProperty] private bool _isOofEnabled;
    [ObservableProperty] private bool _isScheduled;
    [ObservableProperty] private bool _isWorkScheduleEnabled;
    [ObservableProperty] private bool _isAutoSyncEnabled = true;
    // When generating the cloud-sync zip, decide between a managed and an
    // unmanaged Dataverse solution. Managed is the default because Power
    // Automate auto-activates managed flows on import; unmanaged ones land
    // in the Off state and the user has to flip them on manually. Trade-off
    // is that managed flows can't be edited in the cloud designer (the user
    // has to regenerate from this app instead), which matches the
    // "local app is the source of truth" model. Power users who want to
    // tweak the flow online can flip this off.
    [ObservableProperty] private bool _generateManagedCloudSolution = true;
    [ObservableProperty] private bool _isStartWithWindowsEnabled;
    // True when any work-schedule field on the panel has been edited since the
    // last successful Save. Drives the Save button's enabled state and label so
    // the user can tell at a glance whether they have unpersisted changes.
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
    [ObservableProperty] private DateTime _vacationStartDate = DateTime.Today;
    [ObservableProperty] private TimeSpan _vacationStartTime = new(18, 0, 0);
    [ObservableProperty] private DateTime _vacationEndDate = DateTime.Today.AddDays(7);
    [ObservableProperty] private TimeSpan _vacationEndTime = new(9, 0, 0);
    [ObservableProperty] private string _vacationStatus = string.Empty;
    // Persisted across launches; consumed by StartVacation to set
    // DeclineEventsForScheduledOOF on the Scheduled OOF window.
    [ObservableProperty] private bool _declineMeetingsDuringVacation;
    [ObservableProperty] private string _internalReply = string.Empty;
    [ObservableProperty] private string _externalReply = string.Empty;
    [ObservableProperty] private bool _externalAudienceAll = true;
    [ObservableProperty] private DateTime _startDate = DateTime.Today;
    [ObservableProperty] private TimeSpan _startTime = new(9, 0, 0);
    [ObservableProperty] private DateTime _endDate = DateTime.Today.AddDays(1);
    [ObservableProperty] private TimeSpan _endTime = new(9, 0, 0);
    [ObservableProperty] private string _userDisplayName = string.Empty;
    [ObservableProperty] private string _userEmail = string.Empty;
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
        ITemplateService templateService,
        IDialogService dialog,
        INavigationService navigation,
        IPreferencesService prefs,
        ITrayService tray,
        IStartupService startup)
    {
        _exchangeService = exchangeService;
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

    partial void OnIsOofEnabledChanged(bool value)
    {
        if (!value)
        {
            CurrentStatus = OofStatus.Disabled;
            IsScheduled = false;
        }
        else
        {
            CurrentStatus = IsScheduled ? OofStatus.Scheduled : OofStatus.Enabled;
        }
    }

    partial void OnIsScheduledChanged(bool value)
    {
        if (IsOofEnabled)
        {
            CurrentStatus = value ? OofStatus.Scheduled : OofStatus.Enabled;
        }
    }

    partial void OnIsWorkScheduleEnabledChanged(bool value)
    {
        WorkScheduleStatus = value ? GetWorkScheduleStatus() : "Work schedule rule disabled";
        MarkScheduleDirty();
    }

    partial void OnIsAutoSyncEnabledChanged(bool value) => MarkScheduleDirty();
    // Vacation toggle drives both the work-schedule status caption (so the
    // user can see "paused" right on the schedule card) and a refresh of the
    // schedule-status string itself.
    partial void OnIsOnLongVacationChanged(bool value)
    {
        WorkScheduleStatus = GetWorkScheduleStatus();
    }
    // Persist on change so the user's preference survives across launches —
    // suppressed during initial load to avoid a redundant write.
    partial void OnDeclineMeetingsDuringVacationChanged(bool value)
    {
        if (_suppressDirtyTracking) return;
        _prefs.Set("Vacation.DeclineMeetings", value);
    }
    // Cloud-solution flavor toggle isn't part of the schedule, so persist
    // it inline rather than forcing the user to click Save on the schedule
    // card. Suppressed during initial load to avoid a no-op disk write.
    partial void OnGenerateManagedCloudSolutionChanged(bool value)
    {
        if (_suppressDirtyTracking) return;
        _prefs.Set("CloudSync.GenerateManaged", value);
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
        ExternalAudienceAll = template.ExternalAudienceAll;
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

        try
        {
            // Kick off the slow network fetch and the local DB read in parallel
            // immediately — don't gate them on GetCurrentUserAsync (which is just
            // a cached field read but conceptually a precondition).
            var oofTask = _exchangeService.GetOofSettingsAsync();
            var templatesTask = _templateService.GetAllTemplatesAsync();

            var userEmail = await _exchangeService.GetCurrentUserAsync();
            UserDisplayName = userEmail;
            UserEmail = userEmail;
            LoadWorkSchedulePreferences();

            await Task.WhenAll(oofTask, templatesTask);

            var oof = oofTask.Result;
            // The toggle should reflect what Outlook is *actually doing right
            // now*, not the raw mailbox state. A Scheduled OOF whose window is
            // still in the future means OOF replies aren't being sent yet, so
            // we don't want the toggle to read "On". Without this guard, the
            // auto-sync's "next off-hours window" pre-push would flip the
            // toggle on as soon as the user signs in, even mid-workday.
            IsScheduled = oof.Status == OofStatus.Scheduled;
            IsOofEnabled = oof.Status == OofStatus.Enabled
                || (oof.Status == OofStatus.Scheduled
                    && oof.StartTime.HasValue
                    && oof.StartTime.Value <= DateTimeOffset.Now
                    && (!oof.EndTime.HasValue || oof.EndTime.Value > DateTimeOffset.Now));
            // Set CurrentStatus *last* and unconditionally to the real mailbox
            // state. The partial-method side effects of IsOofEnabled /
            // IsScheduled assignments above flip CurrentStatus around as if
            // the user had toggled the switch, which would otherwise lie to
            // the status label about what Exchange actually has.
            CurrentStatus = oof.Status;
            InternalReply = oof.InternalReply;
            ExternalReply = oof.ExternalReply;
            ExternalAudienceAll = oof.ExternalAudienceAll;

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

            Templates.Clear();
            foreach (var t in templatesTask.Result) Templates.Add(t);

            StatusMessage = DescribeOofState(oof);

            _hasLoadedOnce = true;

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
                // banner that won't go away until the next 3-min tick.
                if (HasVacationEnded(DateTime.Now))
                {
                    await EndVacationAsync();
                }
            }
            else if (IsWorkScheduleEnabled)
            {
                // Drop the busy flag *before* triggering the schedule apply +
                // outlook sync. ApplyWorkScheduleAsync and SyncToOutlookCoreAsync
                // both early-return on IsBusy=true, so without this they'd
                // silently no-op on the very first launch — Outlook wouldn't
                // get the initial Scheduled push until the next polling tick.
                IsBusy = false;
                StartAutomationLoop();
                await ApplyWorkScheduleAsync(showSuccessMessage: false, suppressTrayNotification: IsAutoSyncEnabled);
                if (IsAutoSyncEnabled)
                {
                    // Push the current off-hours window straight to Outlook on
                    // launch so the server takes over immediately, even if the
                    // user closes the app right after sign-in.
                    await SyncToOutlookCoreAsync(isUserInitiated: false);
                }
            }

            await MaybePromptStartWithWindowsAsync();
        }
        catch (Exception ex)
        {
            StatusMessage = $"Load failed: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task SaveWorkScheduleAsync()
    {
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
                return;
            }
        }

        SaveWorkSchedulePreferences();
        // Mark prefs as in-sync with the panel; the Save button will go grey
        // again until the user makes another edit.
        HasUnsavedScheduleChanges = false;

        if (IsWorkScheduleEnabled)
        {
            StartAutomationLoop();
            await ApplyWorkScheduleAsync(showSuccessMessage: true, suppressTrayNotification: IsAutoSyncEnabled);
            if (IsAutoSyncEnabled)
            {
                // Schedule edits invalidate any previously-pushed Outlook
                // window, so refresh immediately. The dedupe inside
                // SyncToOutlookCoreAsync already skips if the new window
                // happens to match the cached one.
                _lastSyncedStart = null;
                _lastSyncedEnd = null;
                await SyncToOutlookCoreAsync(isUserInitiated: false);
            }
        }
        else
        {
            StopAutomationLoop();
            StatusMessage = "Work schedule rule saved but not enabled";
            WorkScheduleStatus = "Work schedule rule disabled";
        }
    }

    /// <summary>
    /// "🔄 Sync now" button on the Work Schedule card. Combines what used to
    /// be two separate buttons (Check Now + Sync to Outlook) into one
    /// catch-all force-sync action:
    ///   1. Re-evaluate "are we inside working hours?" right now and flip OOF
    ///      to match (formerly Check Now).
    ///   2. Push the next off-hours window to Outlook so the server keeps
    ///      flipping OOF on its own even if this app isn't running
    ///      (formerly Sync to Outlook).
    /// One button covers ~99% of the cases users actually wanted either of the
    /// old buttons for, while removing the "what's the difference?" question.
    /// </summary>
    [RelayCommand]
    private async Task SyncNowAsync()
    {
        if (!IsWorkScheduleEnabled)
        {
            StatusMessage = "Enable Work Schedule first, then sync.";
            return;
        }

        // 1. Local re-check + Exchange OOF flip.
        await ApplyWorkScheduleAsync(showSuccessMessage: false, suppressTrayNotification: IsAutoSyncEnabled);

        // 2. Force-push the next window to Outlook even if it matches the
        //    cached one — the user explicitly asked for a fresh sync, so the
        //    dedupe inside SyncToOutlookCoreAsync should not skip.
        _lastSyncedStart = null;
        _lastSyncedEnd = null;
        await SyncToOutlookCoreAsync(isUserInitiated: true);
    }

    /// <summary>
    /// Generates a personalised Power Automate setup guide (HTML) that the
    /// user can follow to create a cloud flow which keeps their OOF window
    /// in sync even when every local computer is off. Bound to the
    /// "Generate cloud sync setup guide" button in the Work Schedule card.
    /// </summary>
    [RelayCommand]
    private async Task OpenCloudSyncGuideAsync()
    {
        try
        {
            var snapshot = new WorkScheduleSnapshot(_prefs);
            var path = CloudSyncGuideGenerator.GenerateAndOpen(
                snapshot,
                userEmail: UserEmail,
                internalReply: InternalReply,
                externalReply: ExternalReply,
                externalAudienceAll: ExternalAudienceAll);
            StatusMessage = $"🌐 Cloud sync setup guide opened in your browser ({System.IO.Path.GetFileName(path)})";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to generate cloud sync guide: {ex.Message}";
            await _dialog.AlertAsync("Cloud Sync Guide", ex.Message);
        }
    }

    /// <summary>
    /// Builds a Power Automate Solution .zip the user can upload at
    /// <em>Solutions &rarr; Import solution</em> in make.powerautomate.com.
    /// This gets the user from "I want a cloud flow" to "the flow exists" with
    /// just an upload step instead of the 5-page manual wizard. The zip is
    /// dropped on the desktop so the user can find it easily, and the
    /// containing folder is opened in Explorer with the file pre-selected.
    /// </summary>
    [RelayCommand]
    private async Task GenerateCloudSyncPackageAsync()
    {
        try
        {
            var snapshot = new WorkScheduleSnapshot(_prefs);
            // Desktop is the most discoverable location for a download-style
            // artifact. If we ever ship this on machines with redirected /
            // OneDrive desktops, GetFolderPath returns the redirected path
            // automatically, so this stays correct.
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var outPath = System.IO.Path.Combine(desktop, "OofManager-CloudSync.zip");

            var path = CloudSyncPackageGenerator.Generate(
                snapshot,
                userEmail: UserEmail,
                internalReply: InternalReply,
                externalReply: ExternalReply,
                externalAudienceAll: ExternalAudienceAll,
                generateManaged: GenerateManagedCloudSolution,
                outputPath: outPath);

            // Open the Power Automate Solutions page. We don't pin an env
            // GUID — there's no public way to detect the user's personal
            // environment without pulling in MSAL or the PowerApps PS module,
            // and ~default routes everyone to the tenant Default Environment.
            // The plain /solutions URL lands the user wherever Power Automate
            // last had them (top-right env picker), which is usually right;
            // when it isn't, the user just switches env once via that picker.
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                    "https://make.powerautomate.com/solutions")
                {
                    UseShellExecute = true,
                });
            }
            catch { /* best-effort */ }

            StatusMessage = $"📦 Cloud sync solution saved to {path}. In Power Automate, open Solutions → Import solution and pick this zip.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to generate cloud sync solution: {ex.Message}";
            await _dialog.AlertAsync("Cloud Sync Solution", ex.Message);
        }
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
                    _prefs.Set("Vacation.RestoreExternalAll", ExternalAudienceAll);
                }
            }

            var startOffset = new DateTimeOffset(startLocal, TimeZoneInfo.Local.GetUtcOffset(startLocal));
            var endOffset = new DateTimeOffset(endLocal, TimeZoneInfo.Local.GetUtcOffset(endLocal));

            var settings = new OofSettings
            {
                Status = OofStatus.Scheduled,
                InternalReply = InternalReply,
                ExternalReply = ExternalReply,
                ExternalAudienceAll = ExternalAudienceAll,
                StartTime = startOffset,
                EndTime = endOffset,
                DeclineMeetings = DeclineMeetingsDuringVacation,
            };
            await _exchangeService.SetOofSettingsAsync(settings);

            IsOnLongVacation = true;
            IsScheduled = true;
            var nowOffset = DateTimeOffset.Now;
            IsOofEnabled = startOffset <= nowOffset && endOffset > nowOffset;
            CurrentStatus = OofStatus.Scheduled;

            // Invalidate any work-schedule sync cache so when vacation ends
            // the next normal sync re-pushes the regular off-hours window.
            _lastSyncedStart = null;
            _lastSyncedEnd = null;

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
        var restoreAll = _prefs.GetBool("Vacation.RestoreExternalAll", true);

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
                    ExternalAudienceAll = restoreAll,
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
                    ExternalAudienceAll = restoreAll,
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
                ExternalAudienceAll = restoreAll,
            };
        }

        IsBusy = true;
        try
        {
            await _exchangeService.SetOofSettingsAsync(target);
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

        InternalReply = restoreInternal;
        ExternalReply = restoreExternal;
        ExternalAudienceAll = restoreAll;

        // Mirror the pushed state into the local UI flags. Order matches the
        // post-Set bookkeeping in SyncToOutlookCoreAsync / ApplyWorkScheduleAsync
        // so dedupe caches stay in step with what the server now has.
        if (target.Status == OofStatus.Scheduled)
        {
            IsScheduled = true;
            var nowOffset = DateTimeOffset.Now;
            IsOofEnabled = target.StartTime!.Value <= nowOffset && target.EndTime!.Value > nowOffset;
            CurrentStatus = OofStatus.Scheduled;
            _lastSyncedStart = target.StartTime;
            _lastSyncedEnd = target.EndTime;
        }
        else
        {
            IsScheduled = false;
            IsOofEnabled = false;
            CurrentStatus = OofStatus.Disabled;
            _lastSyncedStart = null;
            _lastSyncedEnd = null;
        }

        VacationStatus = string.Empty;
        StatusMessage = "🛬 Vacation ended — OOF cleared.";

        if (_tray.IsWindowHidden)
        {
            _tray.ShowNotification(
                "OOF Manager",
                "Vacation ended — out-of-office cleared, regular work schedule resumed.");
        }

        // If Work Schedule is off and we only kept the loop alive for vacation
        // end-detection, shut it back down so we don't poll for nothing.
        if (!IsWorkScheduleEnabled)
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
            return;
        }
        var (start, end) = window.Value;

        // Dedupe: skip the server round-trip if the window we'd push is exactly
        // the one we last pushed. The auto-sync loop fires every few minutes,
        // so without this we'd hammer Exchange with identical Set calls all
        // through a single off-hours stretch.
        if (!isUserInitiated
            && _lastSyncedStart == start
            && _lastSyncedEnd == end)
        {
            return;
        }

        IsBusy = true;
        try
        {
            var settings = new OofSettings
            {
                Status = OofStatus.Scheduled,
                InternalReply = InternalReply,
                ExternalReply = ExternalReply,
                ExternalAudienceAll = ExternalAudienceAll,
                StartTime = start,
                EndTime = end
            };

            await _exchangeService.SetOofSettingsAsync(settings);

            _lastSyncedStart = start;
            _lastSyncedEnd = end;

            // Reflect the new server state locally so the UI's "Out of Office"
            // card matches Outlook (which is now in Scheduled mode). The
            // toggle reflects whether OOF is *actually firing right now* — a
            // future-only Scheduled window means it isn't, so the toggle stays
            // Off until the start time arrives. Without this guard the toggle
            // would visually flip On the moment the user clicks Sync now mid-
            // workday, which lies about whether senders are actually getting
            // auto-replies.
            IsScheduled = true;
            var nowOffset = DateTimeOffset.Now;
            IsOofEnabled = start <= nowOffset && end > nowOffset;
            // Re-assert CurrentStatus *after* the IsOofEnabled / IsScheduled
            // setters' partial methods run, so the status label always reads
            // the real mailbox state instead of being clobbered to Disabled
            // when IsOofEnabled is false.
            CurrentStatus = OofStatus.Scheduled;

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
            // 3-min self-heal loop doesn't spam balloons every tick — the user
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

    private async Task ApplyWorkScheduleAsync(bool showSuccessMessage, bool suppressTrayNotification = false)
    {
        if (!IsWorkScheduleEnabled || IsBusy) return;
        // Long-vacation mode owns the OOF window outright; the regular work
        // schedule must not flip OOF on/off underneath it or we'd erase the
        // multi-day Scheduled block the user explicitly created.
        if (IsOnLongVacation) return;

        var shouldBeOof = !IsNowInsideWorkingHours(DateTime.Now);
        var targetStatus = shouldBeOof ? OofStatus.Enabled : OofStatus.Disabled;
        var previousStatus = CurrentStatus;

        // We deliberately do NOT short-circuit when CurrentStatus == targetStatus.
        // The local CurrentStatus can drift from the real server state (cached
        // value at login, change made from another device, change made by an
        // admin, etc.). Always pushing the desired state guarantees the local
        // UI's claim ("OOF off"/"OOF on") matches what the tenant actually
        // has — and SetOofSettingsAsync now reads back from Exchange and throws
        // if the change didn't take effect, so any silent failures surface.

        IsBusy = true;
        try
        {
            var settings = new OofSettings
            {
                Status = targetStatus,
                InternalReply = InternalReply,
                ExternalReply = ExternalReply,
                ExternalAudienceAll = ExternalAudienceAll
            };

            await _exchangeService.SetOofSettingsAsync(settings);
            CurrentStatus = targetStatus;
            IsOofEnabled = shouldBeOof;
            // We just overwrote any Scheduled window on Exchange with a flat
            // Enabled/Disabled state, so the dedupe cache no longer reflects
            // server reality. Without this invalidation the auto-sync loop
            // (ApplyWorkScheduleAsync → SyncToOutlookCoreAsync) leaves OOF
            // sitting at Disabled all through working hours, because the
            // computed off-hours window matches the stale cache and the
            // follow-up push gets skipped — which is exactly why "auto-sync
            // turns OOF off" while "manual Sync now schedules the window".
            _lastSyncedStart = null;
            _lastSyncedEnd = null;
            // Re-raise PropertyChanged so any binding that already cached the
            // value (e.g. after the OnIsOofEnabledChanged partial method ran
            // and re-assigned CurrentStatus to the same OofStatus) still gets
            // a refresh notification. Cheap and idempotent.
            OnPropertyChanged(nameof(CurrentStatus));
            OnPropertyChanged(nameof(IsOofEnabled));
            WorkScheduleStatus = GetWorkScheduleStatus();
            // Always surface the auto-switch in the status bar so the user
            // can see exactly when the schedule kicked in, not only on the
            // “Save and Apply” / “Check Now” code path.
            var nowLabel = DateTime.Now.ToString("HH:mm");
            StatusMessage = shouldBeOof
                ? $"✅ {nowLabel} · Outside working hours: OOF turned on automatically"
                : $"✅ {nowLabel} · Inside working hours: OOF turned off automatically";

            // Tray notification: only fire on a *real* flip and only when the
            // window is hidden — otherwise the status bar already tells the
            // user, and a balloon for every 5-minute self-heal would be spam.
            // Also suppress when the caller will immediately push a Scheduled
            // window via SyncToOutlookCoreAsync — that path balloons the more
            // informative "auto-reply scheduled: <window>" message, and the
            // intermediate "OOF turned OFF" balloon is misleading because the
            // flat Disabled state we just pushed is about to be replaced by a
            // Scheduled state in the same tick.
            if (previousStatus != targetStatus && _tray.IsWindowHidden && !suppressTrayNotification)
            {
                _tray.ShowNotification(
                    "OOF Manager",
                    shouldBeOof
                        ? $"Out-of-office turned ON ({nowLabel}) — outside working hours"
                        : $"Out-of-office turned OFF ({nowLabel}) — inside working hours");
            }
        }
        catch (Exception ex)
        {
            // Always surface failures, even from the silent background loop —
            // a silent failure is exactly what hid the original "can't change
            // OOF" bug.
            StatusMessage = $"Failed to apply work schedule rule: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task SaveOofAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        StatusMessage = "Saving OOF settings...";

        try
        {
            var settings = new OofSettings
            {
                Status = CurrentStatus,
                InternalReply = InternalReply,
                ExternalReply = ExternalReply,
                ExternalAudienceAll = ExternalAudienceAll,
            };

            if (CurrentStatus == OofStatus.Scheduled)
            {
                // Compute the UTC offset *at the scheduled instant*, not at save-time.
                // Otherwise a window that straddles DST would be off by one hour.
                var startLocal = StartDate.Add(StartTime);
                var endLocal = EndDate.Add(EndTime);
                settings.StartTime = new DateTimeOffset(startLocal, TimeZoneInfo.Local.GetUtcOffset(startLocal));
                settings.EndTime = new DateTimeOffset(endLocal, TimeZoneInfo.Local.GetUtcOffset(endLocal));
            }

            await _exchangeService.SetOofSettingsAsync(settings);
            StatusMessage = CurrentStatus == OofStatus.Disabled
                ? "✅ OOF turned off"
                : "✅ OOF saved and enabled";
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
    private async Task QuickToggleAsync()
    {
        // The Out-of-Office card already has a ToggleSwitch bound to IsOofEnabled,
        // so this button only needs to *persist* whatever the user picked. The
        // previous implementation flipped IsOofEnabled before saving, which double-
        // toggled when the user had already moved the switch (switch ON + click =
        // server gets OFF), making it look like saves "didn't take".

        // Normalize away any leftover Scheduled state from a previous auto-sync
        // before saving. Without this, IsScheduled stays true after LoadAsync
        // pulls a Scheduled OOF from Exchange, so Save Now would re-push the
        // stale 17:30→09:00 window and Outlook keeps displaying it as a
        // schedule even though the user is in manual mode.
        if (IsScheduled)
        {
            IsScheduled = false;
            // Re-derive CurrentStatus from the toggle now that IsScheduled is
            // false (Enabled if on, Disabled if off).
            CurrentStatus = IsOofEnabled ? OofStatus.Enabled : OofStatus.Disabled;
        }
        await SaveOofAsync();
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
            ExternalAudienceAll = ExternalAudienceAll
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

    /// <summary>
    /// Builds the human-readable status line shown at the top of the main
    /// page. Three real cases the user cares about:
    ///   - Disabled                : OOF is off entirely
    ///   - Enabled (no schedule)   : OOF is on permanently / until manually
    ///                               turned off — no end time set on the mailbox
    ///   - Scheduled               : OOF will only be active inside the
    ///                               start→end window. Further split into
    ///                               "active now" vs. "scheduled to start later"
    ///                               so the user can tell at a glance whether
    ///                               replies are actually going out right now.
    /// </summary>
    private static string DescribeOofState(OofSettings oof)
    {
        switch (oof.Status)
        {
            case OofStatus.Disabled:
                return "🔕 OOF is currently OFF";

            case OofStatus.Enabled:
                // No end time = "until I turn it off myself". Call that out
                // explicitly so the user doesn't expect Exchange to disable it
                // automatically.
                return "🟢 OOF is ON (no schedule — stays on until manually turned off)";

            case OofStatus.Scheduled:
                var now = DateTimeOffset.Now;
                var startStr = oof.StartTime?.LocalDateTime.ToString("ddd MM-dd HH:mm");
                var endStr = oof.EndTime?.LocalDateTime.ToString("ddd MM-dd HH:mm");

                // If both ends are present we can tell the user *which side*
                // of the window we're currently sitting in.
                if (oof.StartTime.HasValue && oof.EndTime.HasValue)
                {
                    if (now < oof.StartTime.Value)
                        return $"🟡 OOF is scheduled — will start {startStr} and end {endStr}";
                    if (now >= oof.EndTime.Value)
                        return $"🔕 OOF schedule already ended ({endStr}) — currently OFF";
                    return $"🟢 OOF is ON (scheduled until {endStr})";
                }

                // Defensive fallback: tenant returned Scheduled but no times.
                return "🟡 OOF is scheduled";

            default:
                return string.Empty;
        }
    }

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
            IsWorkScheduleEnabled = _prefs.GetBool("WorkSchedule.Enabled", false);
        IsAutoSyncEnabled = _prefs.GetBool("WorkSchedule.AutoSync", true);
        GenerateManagedCloudSolution = _prefs.GetBool("CloudSync.GenerateManaged", true);
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
        DeclineMeetingsDuringVacation = _prefs.GetBool("Vacation.DeclineMeetings", false);
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
            _prefs.Set("WorkSchedule.AutoSync", IsAutoSyncEnabled);
            _prefs.Set("CloudSync.GenerateManaged", GenerateManagedCloudSolution);
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

    /// <summary>
    /// Asks the user once (per profile) whether they want OofManager to launch
    /// at Windows logon and start hidden in the tray. Designed to fire after
    /// the *first* successful sign-in so the user has already seen the app
    /// works before being asked to bake it into their startup.
    /// </summary>
    private async Task MaybePromptStartWithWindowsAsync()
    {
        // Already enabled (e.g. via Task Manager > Startup, a previous session,
        // or auto-healed from a stale Run entry)? Just sync the toggle and
        // skip the prompt — and crucially, don't burn the "shown" flag, so
        // that if the user later disables it externally we still get a chance
        // to prompt on a future launch.
        if (_startup.IsEnabled)
        {
            IsStartWithWindowsEnabled = true;
            return;
        }

        // Already asked once on this profile and the user said no — respect that.
        if (_startup.HasBeenPromptedBefore()) return;

        var enable = await _dialog.ConfirmAsync(
            "Start with Windows?",
            "Would you like OofManager to launch automatically when you sign in to Windows and stay hidden in the tray?\n\n"
            + "This lets it keep your Outlook OOF in sync without you remembering to open it.\n\n"
            + "You can change this any time from the main page.",
            accept: "Yes, enable",
            cancel: "No thanks");

        // Only record the prompt as "shown" once it actually appeared, so a
        // crash / early-exit before the dialog renders doesn't permanently
        // suppress the question.
        _startup.MarkPromptShown();

        if (enable)
        {
            IsStartWithWindowsEnabled = true; // setter writes the registry
            StatusMessage = "🚀 OofManager will start with Windows and run in the tray.";
        }
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
                // Sleep until either the next work-hours boundary or 3 minutes
                // (whichever is sooner). The boundary-aligned wake guarantees
                // we flip OOF *exactly* at start/end times, while the 3-minute
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
                    // Vacation owns the OOF window while it's on. The only
                    // thing we do here is detect the end-of-vacation moment
                    // and hand control back to the regular work schedule.
                    if (IsOnLongVacation)
                    {
                        if (HasVacationEnded(DateTime.Now))
                        {
                            await EndVacationAsync();
                        }
                        return;
                    }
                    await ApplyWorkScheduleAsync(showSuccessMessage: false, suppressTrayNotification: IsAutoSyncEnabled);
                    if (IsAutoSyncEnabled)
                    {
                        // Auto-sync runs after the local toggle so Outlook stays
                        // in lock-step with the client's view, and so the
                        // *next* off-hours window is already pre-pushed before
                        // the user's working day even ends.
                        await SyncToOutlookCoreAsync(isUserInitiated: false);
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
        // Cap the wait to 3 minutes so we self-heal if the clock changes,
        // workdays change, etc.
        var fallback = TimeSpan.FromMinutes(3);
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
}
