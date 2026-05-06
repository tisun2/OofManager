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
    [ObservableProperty] private bool _isSaturdayWorkday;
    [ObservableProperty] private bool _isSundayWorkday;
    [ObservableProperty] private TimeSpan _workStartTime = new(9, 0, 0);
    [ObservableProperty] private TimeSpan _workEndTime = new(18, 0, 0);
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
    public IReadOnlyList<TimeSpan> WorkTimeOptions { get; } =
        Enumerable.Range(0, 48).Select(i => TimeSpan.FromMinutes(i * 30)).ToList();

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
        if (WorkEndTime <= WorkStartTime)
        {
            await _dialog.AlertAsync("Invalid Work Hours", "End time must be later than start time.");
            return;
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
            WorkScheduleStatus = GetWorkScheduleStatus();
            if (showSuccessMessage)
            {
                StatusMessage = shouldBeOof
                    ? "✅ Outside working hours: OOF turned on automatically"
                    : "✅ Inside working hours: OOF turned off automatically";
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
        return now.TimeOfDay >= WorkStartTime && now.TimeOfDay < WorkEndTime;
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
        IsSaturdayWorkday = _prefs.GetBool("WorkSchedule.Saturday", false);
        IsSundayWorkday = _prefs.GetBool("WorkSchedule.Sunday", false);
        // Snap the persisted values to the same 30-minute grid the dropdown
        // exposes; otherwise an off-grid value (from an older build that
        // accepted free text) would make the ComboBox show blank.
        WorkStartTime = SnapToHalfHour(TimeSpan.FromMinutes(_prefs.GetInt("WorkSchedule.StartMinutes", 9 * 60)));
        WorkEndTime = SnapToHalfHour(TimeSpan.FromMinutes(_prefs.GetInt("WorkSchedule.EndMinutes", 18 * 60)));
        WorkScheduleStatus = GetWorkScheduleStatus();
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
            _prefs.Set("WorkSchedule.StartMinutes", (int)WorkStartTime.TotalMinutes);
            _prefs.Set("WorkSchedule.EndMinutes", (int)WorkEndTime.TotalMinutes);
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
        var startToday = now.Date.Add(WorkStartTime);
        var endToday = now.Date.Add(WorkEndTime);
        TimeSpan? candidate = null;
        if (now < startToday) candidate = startToday - now;
        else if (now < endToday) candidate = endToday - now;
        else candidate = (now.Date.AddDays(1).Add(WorkStartTime)) - now;

        var delay = candidate.HasValue && candidate.Value < fallback ? candidate.Value : fallback;
        // Don't poll faster than once a minute regardless.
        return delay < TimeSpan.FromMinutes(1) ? TimeSpan.FromMinutes(1) : delay;
    }
}
