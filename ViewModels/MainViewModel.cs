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
    private CancellationTokenSource? _automationCts;
    private bool _hasLoadedOnce;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private OofStatus _currentStatus = OofStatus.Disabled;
    [ObservableProperty] private bool _isOofEnabled;
    [ObservableProperty] private bool _isScheduled;
    [ObservableProperty] private bool _isWorkScheduleEnabled;
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
        ITrayService tray)
    {
        _exchangeService = exchangeService;
        _templateService = templateService;
        _dialog = dialog;
        _navigation = navigation;
        _prefs = prefs;
        _tray = tray;
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
            CurrentStatus = oof.Status;
            IsOofEnabled = oof.Status != OofStatus.Disabled;
            IsScheduled = oof.Status == OofStatus.Scheduled;
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

            StatusMessage = oof.Status == OofStatus.Disabled
                ? "OOF is currently off"
                : "OOF is currently on";

            _hasLoadedOnce = true;

            if (IsWorkScheduleEnabled)
            {
                StartAutomationLoop();
                await ApplyWorkScheduleAsync(showSuccessMessage: false);
            }
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

        if (IsWorkScheduleEnabled)
        {
            StartAutomationLoop();
            await ApplyWorkScheduleAsync(showSuccessMessage: true);
        }
        else
        {
            StopAutomationLoop();
            StatusMessage = "Work schedule rule saved but not enabled";
            WorkScheduleStatus = "Work schedule rule disabled";
        }
    }

    [RelayCommand]
    private async Task ApplyWorkScheduleAsync()
    {
        await ApplyWorkScheduleAsync(showSuccessMessage: true);
    }

    private async Task ApplyWorkScheduleAsync(bool showSuccessMessage)
    {
        if (!IsWorkScheduleEnabled || IsBusy) return;

        var shouldBeOof = !IsNowInsideWorkingHours(DateTime.Now);
        var targetStatus = shouldBeOof ? OofStatus.Enabled : OofStatus.Disabled;

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
    private async Task LogoutAsync()
    {
        StopAutomationLoop();
        // Reset the load gate so a re-login refreshes data instead of showing stale state.
        _hasLoadedOnce = false;
        await _exchangeService.DisconnectAsync();
        _navigation.NavigateToLogin();
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
        return IsNowInsideWorkingHours(DateTime.Now)
            ? "Currently inside working hours: OOF should be off"
            : "Currently outside working hours: OOF should be on";
    }

    private void LoadWorkSchedulePreferences()
    {
        IsWorkScheduleEnabled = _prefs.GetBool("WorkSchedule.Enabled", false);
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

        WorkScheduleStatus = GetWorkScheduleStatus();
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
                var op = Application.Current.Dispatcher.InvokeAsync(
                    () => ApplyWorkScheduleAsync(showSuccessMessage: false));
                await op.Task.Unwrap();
            }
        }
        catch (OperationCanceledException) { }
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
}
