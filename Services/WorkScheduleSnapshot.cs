namespace OofManager.Wpf.Services;

/// <summary>
/// Immutable snapshot of the user's per-day work schedule, loaded from
/// <see cref="IPreferencesService"/>. Pure logic, no UI dependencies, so it can
/// be used both from <c>MainViewModel</c> (interactive) and from the
/// <c>--sync</c> background runner (no window).
/// </summary>
public sealed class WorkScheduleSnapshot
{
    public bool IsEnabled { get; }
    public bool BackgroundSyncEnabled { get; }
    private readonly bool[] _isWorkday = new bool[7]; // index = (int)DayOfWeek
    private readonly TimeSpan[] _start = new TimeSpan[7];
    private readonly TimeSpan[] _end = new TimeSpan[7];

    public WorkScheduleSnapshot(IPreferencesService prefs)
    {
        IsEnabled = prefs.GetBool("WorkSchedule.Enabled", false);
        BackgroundSyncEnabled = prefs.GetBool("WorkSchedule.BackgroundSync", false);

        var legacyStart = prefs.GetInt("WorkSchedule.StartMinutes", 9 * 60);
        var legacyEnd = prefs.GetInt("WorkSchedule.EndMinutes", 18 * 60);

        foreach (var day in new[]
        {
            DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
            DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday
        })
        {
            var name = day.ToString();
            var defaultWorkday = day != DayOfWeek.Saturday && day != DayOfWeek.Sunday;
            // Saved keys originally defaulted Sat/Sun to true; preserve that
            // behaviour here so an existing user's stored value wins. (The
            // legacy default in MainViewModel's LoadWorkSchedulePreferences
            // is also `true` for every day, so we mirror that.)
            _isWorkday[(int)day] = prefs.GetBool($"WorkSchedule.{name}", true);
            _start[(int)day] = TimeSpan.FromMinutes(prefs.GetInt($"WorkSchedule.{name}.StartMinutes", legacyStart));
            _end[(int)day] = TimeSpan.FromMinutes(prefs.GetInt($"WorkSchedule.{name}.EndMinutes", legacyEnd));
        }
    }

    public bool IsWorkday(DayOfWeek day) => _isWorkday[(int)day];
    public TimeSpan GetStart(DayOfWeek day) => _start[(int)day];
    public TimeSpan GetEnd(DayOfWeek day) => _end[(int)day];

    /// <summary>
    /// True when <paramref name="now"/> falls inside the configured working
    /// hours for that day-of-week. Off-work days always return false.
    /// </summary>
    public bool IsNowInsideWorkingHours(DateTime now)
    {
        if (!IsWorkday(now.DayOfWeek)) return false;
        var t = now.TimeOfDay;
        return t >= GetStart(now.DayOfWeek) && t < GetEnd(now.DayOfWeek);
    }

    /// <summary>
    /// Next contiguous off-hours window starting at or after <paramref name="now"/>.
    /// Mirrors <c>MainViewModel.ComputeNextOffHoursWindow</c>: anchors the start
    /// to today's end-of-work whenever we're past it (so an evening run pushes
    /// "tonight's" window, not a window starting at the random run time).
    /// Returns null when no work day is configured (off-hours has no end anchor).
    /// </summary>
    public (DateTimeOffset start, DateTimeOffset end)? ComputeNextOffHoursWindow(DateTime now)
    {
        DateTime offStart;
        if (IsWorkday(now.DayOfWeek))
        {
            var startToday = now.Date.Add(GetStart(now.DayOfWeek));
            var endToday = now.Date.Add(GetEnd(now.DayOfWeek));
            // Pre-work: walk back to the previous workday's end-of-work.
            // Inside/after work: anchor to today's end-of-work.
            offStart = now < startToday
                ? (FindMostRecentEndOfWorkAtOrBefore(now) ?? now)
                : endToday;
        }
        else
        {
            // Off-work day (e.g. Saturday): anchor to the last workday's
            // end-of-work, falling back to "now" only when no workday is
            // configured at all.
            offStart = FindMostRecentEndOfWorkAtOrBefore(now) ?? now;
        }

        var offEnd = FindNextStartOfWorkAfter(offStart);
        if (offEnd == null || offEnd.Value <= offStart) return null;

        return (
            new DateTimeOffset(offStart, TimeZoneInfo.Local.GetUtcOffset(offStart)),
            new DateTimeOffset(offEnd.Value, TimeZoneInfo.Local.GetUtcOffset(offEnd.Value))
        );
    }

    private DateTime? FindNextStartOfWorkAfter(DateTime t)
    {
        if (IsWorkday(t.DayOfWeek))
        {
            var startToday = t.Date.Add(GetStart(t.DayOfWeek));
            if (t < startToday) return startToday;
        }
        for (int i = 1; i <= 7; i++)
        {
            var nextDate = t.Date.AddDays(i);
            if (IsWorkday(nextDate.DayOfWeek))
                return nextDate.Add(GetStart(nextDate.DayOfWeek));
        }
        return null;
    }

    private DateTime? FindMostRecentEndOfWorkAtOrBefore(DateTime t)
    {
        if (IsWorkday(t.DayOfWeek))
        {
            var endToday = t.Date.Add(GetEnd(t.DayOfWeek));
            if (t >= endToday) return endToday;
        }
        for (int i = 1; i <= 7; i++)
        {
            var prevDate = t.Date.AddDays(-i);
            if (IsWorkday(prevDate.DayOfWeek))
                return prevDate.Add(GetEnd(prevDate.DayOfWeek));
        }
        return null;
    }
}
