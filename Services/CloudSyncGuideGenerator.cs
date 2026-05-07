using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using OofManager.Wpf.Models;

namespace OofManager.Wpf.Services;

/// <summary>
/// Generates a personalised HTML guide that walks the user through creating a
/// Power Automate scheduled cloud flow which mirrors their local OofManager
/// work-schedule rules. The flow runs in Microsoft 365 (no local computer
/// required), so OOF replies keep being scheduled in Outlook even when the
/// user's machines are powered off — solving the "computer must be on" gap
/// in the local Scheduled-Task background sync.
///
/// Output is a single self-contained .html file with inline CSS and clipboard
/// helpers, dropped into <c>%TEMP%</c> and opened in the default browser. We
/// deliberately do NOT touch Power Automate APIs from the app: that would
/// require the user to grant a separate consent surface, and the import-zip
/// path needs Power Apps connection-reference fix-ups that aren't worth the
/// brittleness for a setup-once feature.
/// </summary>
public static class CloudSyncGuideGenerator
{
    /// <summary>
    /// Builds the guide for the given snapshot + reply settings, writes it to
    /// a stable filename in <c>%TEMP%</c>, and launches it in the default
    /// browser. Returns the file path so the caller can also surface it in
    /// the status bar.
    /// </summary>
    public static string GenerateAndOpen(
        WorkScheduleSnapshot schedule,
        string userEmail,
        string internalReply,
        string externalReply,
        bool externalAudienceAll)
    {
        var html = Build(schedule, userEmail, internalReply, externalReply, externalAudienceAll);

        var path = Path.Combine(Path.GetTempPath(), "OofManager-CloudSync-Setup.html");
        File.WriteAllText(path, html, Encoding.UTF8);

        try
        {
            // UseShellExecute=true so it opens in whatever the user has set
            // as the default browser, not as a child of OofManager.exe (which
            // would block until the browser closes on some configurations).
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
            // Non-fatal — the caller can show the path in the status bar so
            // the user can open it themselves.
        }

        return path;
    }

    private static string Build(
        WorkScheduleSnapshot schedule,
        string userEmail,
        string internalReply,
        string externalReply,
        bool externalAudienceAll)
    {
        // Power Automate's recurrence uses the M365 account's time zone, but
        // expression helpers like utcNow() return UTC. Without an explicit
        // convertFromUtc we'd compute "tomorrow 09:00 UTC" instead of "tomorrow
        // 09:00 LOCAL", which is several hours off in non-UTC time zones. Bake
        // the local Windows TZ id into both expressions and the Time Zone
        // field of the Set automatic replies action so the user doesn't have
        // to think about it.
        var tzId = TimeZoneInfo.Local.Id;
        var tzDisplay = TimeZoneInfo.Local.DisplayName;

        // Pick a representative work-end / work-start time for the recurrence.
        // We use the EARLIEST end and LATEST start across all configured
        // workdays, which gives a conservative "OOF should be on" window
        // covering the whole week. Users with significant per-day variation
        // see a "Customize" callout in the rendered guide.
        var workEnd = ComputeRepresentativeEnd(schedule);
        var workStart = ComputeRepresentativeStart(schedule);

        // Recurrence fires daily at this hour:minute (local time of the
        // signed-in M365 account, which Power Automate uses by default).
        var triggerHour = workEnd.Hours;
        var triggerMinute = workEnd.Minutes;

        // The flow body needs an End time that lands at the next workday's
        // start-of-work. For Mon-Thu that's tomorrow; for Friday it's three
        // days later. We bake this rule into the dynamic content via a
        // Power Automate expression that computes hours-until-next-workstart.
        var endHour = workStart.Hours;
        var endMinute = workStart.Minutes;

        // For each weekday, decide how many days to add to "today" to land on
        // the next configured workday. Used in the if/switch expression we
        // emit into the End Time field.
        var hopDays = new int[7];
        for (int i = 0; i < 7; i++)
        {
            var day = (DayOfWeek)i;
            // From this day, walk forward until we hit a workday.
            for (int hop = 1; hop <= 7; hop++)
            {
                var candidate = (DayOfWeek)(((int)day + hop) % 7);
                if (schedule.IsWorkday(candidate))
                {
                    hopDays[i] = hop;
                    break;
                }
            }
        }

        // The Power Automate expression switches on dayOfWeek(utcNow())
        // (Sun=0..Sat=6) and adds the right number of days to today's start,
        // then offsets to the work-start hour:minute. Wrapping in
        // formatDateTime keeps the field happy as an ISO string.
        var endTimeExpression = BuildEndTimeExpression(hopDays, endHour, endMinute, tzId);
        var startTimeExpression = BuildStartTimeExpression(tzId);

        var internalEsc = Escape(internalReply);
        var externalEsc = Escape(externalReply);
        var userEsc = Escape(userEmail);
        var audienceLabel = externalAudienceAll ? "All" : "Known (contacts only)";
        var weeklySummary = BuildWeeklyScheduleSummary(schedule);
        var hasPerDayVariation = HasPerDayVariation(schedule);

        var sb = new StringBuilder(8192);
        sb.Append(@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""utf-8""/>
<title>OofManager — Cloud Sync Setup Guide</title>
<style>
  :root { --accent: #0078D4; --accent-dark: #005A9E; --bg: #FAFAFA; --card: #fff; --border: #E5E5E5; --muted: #666; --code-bg: #F4F4F4; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, ""Segoe UI"", Roboto, sans-serif; background: var(--bg); color: #1A1A1A; margin: 0; padding: 0; line-height: 1.55; }
  .container { max-width: 860px; margin: 0 auto; padding: 32px 24px 64px; }
  h1 { font-size: 28px; margin: 0 0 8px; }
  h2 { font-size: 20px; margin: 32px 0 12px; padding-bottom: 8px; border-bottom: 2px solid var(--accent); }
  h3 { font-size: 16px; margin: 20px 0 8px; color: var(--accent-dark); }
  .lede { color: var(--muted); margin: 0 0 24px; font-size: 15px; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 20px 24px; margin: 16px 0; }
  .step { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 20px 24px; margin: 16px 0; position: relative; }
  .step .num { position: absolute; left: -16px; top: 18px; background: var(--accent); color: #fff; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.15); }
  .step h3 { margin-top: 0; }
  code { background: var(--code-bg); padding: 2px 6px; border-radius: 3px; font-family: ""Cascadia Code"", Consolas, monospace; font-size: 13px; }
  pre { background: var(--code-bg); border: 1px solid var(--border); padding: 12px 16px; border-radius: 6px; overflow-x: auto; font-family: ""Cascadia Code"", Consolas, monospace; font-size: 13px; line-height: 1.45; margin: 8px 0; white-space: pre-wrap; word-break: break-word; }
  .copy-row { display: flex; gap: 8px; align-items: flex-start; }
  .copy-row pre { flex: 1; margin: 0; }
  button.copy { background: var(--accent); color: #fff; border: none; padding: 8px 14px; border-radius: 4px; cursor: pointer; font-size: 13px; white-space: nowrap; min-width: 64px; }
  button.copy:hover { background: var(--accent-dark); }
  button.copy.copied { background: #107C10; }
  a { color: var(--accent); }
  .callout { background: #FFF8E1; border-left: 4px solid #F4B400; padding: 12px 16px; border-radius: 4px; margin: 12px 0; font-size: 14px; }
  .callout.info { background: #E7F3FE; border-color: var(--accent); }
  .callout.success { background: #E8F5E9; border-color: #107C10; }
  .field { display: grid; grid-template-columns: 180px 1fr; gap: 8px 16px; margin: 8px 0; align-items: center; font-size: 14px; }
  .field .label { color: var(--muted); font-weight: 500; }
  .field .value code { display: inline-block; }
  table { border-collapse: collapse; margin: 12px 0; font-size: 14px; }
  th, td { padding: 6px 12px; border: 1px solid var(--border); text-align: left; }
  th { background: var(--code-bg); }
  .hero-button { display: inline-block; background: var(--accent); color: #fff !important; padding: 12px 22px; text-decoration: none; border-radius: 6px; font-weight: 600; margin: 8px 0; }
  .hero-button:hover { background: var(--accent-dark); }
  .footer { color: var(--muted); font-size: 12px; margin-top: 48px; text-align: center; }
</style>
</head>
<body>
<div class=""container"">
  <h1>🌐 OofManager &mdash; Cloud Sync Setup</h1>
  <p class=""lede"">Set up a Power Automate scheduled cloud flow so your Out-of-Office window is pushed to Outlook every morning automatically &mdash; even when all of your computers are powered off. Takes about 5 minutes.</p>

  <div class=""card"">
    <h3 style=""margin-top:0"">Why do I need this?</h3>
    <p>OofManager already pushes your next OOF window to Outlook when its app or background scheduled task is running. But if every machine you use is fully powered off (or you're travelling without your laptop), nothing on your end can talk to Exchange. A cloud flow runs inside Microsoft 365 itself, so it keeps working regardless of what your devices are doing.</p>
  </div>

  <h2>Your personalised settings</h2>
  <div class=""card"">
    <div class=""field""><span class=""label"">Account</span><span class=""value""><code>");
        sb.Append(userEsc);
        sb.Append(@"</code></span></div>
    <div class=""field""><span class=""label"">Trigger time (daily)</span><span class=""value""><code>");
        sb.Append(triggerHour.ToString("00") + ":" + triggerMinute.ToString("00"));
        sb.Append(@"</code> &middot; runs every workday at end of work</span></div>
    <div class=""field""><span class=""label"">External audience</span><span class=""value""><code>");
        sb.Append(audienceLabel);
        sb.Append(@"</code></span></div>
    <div class=""field""><span class=""label"">Weekly schedule</span><span class=""value"">");
        sb.Append(weeklySummary);
        sb.Append(@"</span></div>
  </div>");

        if (hasPerDayVariation)
        {
            sb.Append(@"
  <div class=""callout"">
    <strong>Heads up:</strong> Your work hours vary by day. The default flow uses your earliest end-of-work time as the trigger and your latest start-of-work time as the resume time, which gives a conservative window that covers every workday. If you want exact per-day boundaries, see the <em>Customize per day</em> section near the bottom of this page.
  </div>");
        }

        sb.Append(@"

  <h2>Step-by-step</h2>

  <div class=""step""><span class=""num"">1</span>
    <h3>Open Power Automate and start a scheduled flow</h3>
    <p>Sign in with the same Microsoft 365 account you use for Outlook (<code>");
        sb.Append(userEsc);
        sb.Append(@"</code>).</p>
    <a class=""hero-button"" href=""https://make.powerautomate.com/manage/environments/~default/flows/new"" target=""_blank"" rel=""noopener"">Open Power Automate &rarr;</a>
    <p>Click <strong>Create</strong> in the left rail, then choose <strong>Scheduled cloud flow</strong>.</p>
  </div>

  <div class=""step""><span class=""num"">2</span>
    <h3>Configure the recurrence trigger</h3>
    <p>Use these exact values:</p>
    <table>
      <tr><th>Field</th><th>Value</th></tr>
      <tr><td>Flow name</td><td><code>OofManager Cloud Sync</code></td></tr>
      <tr><td>Starting</td><td>(leave today's date, choose a time slightly after now)</td></tr>
      <tr><td>Repeat every</td><td><code>1 Day</code></td></tr>
      <tr><td>At these hours</td><td><code>");
        sb.Append(triggerHour.ToString("00"));
        sb.Append(@"</code></td></tr>
      <tr><td>At these minutes</td><td><code>");
        sb.Append(triggerMinute.ToString("00"));
        sb.Append(@"</code></td></tr>
      <tr><td>On these days</td><td>");
        sb.Append(BuildDaysSelectionLabel(schedule));
        sb.Append(@"</td></tr>
    </table>
    <p>Click <strong>Create</strong>.</p>
  </div>

  <div class=""step""><span class=""num"">3</span>
    <h3>Add the &ldquo;Set up automatic replies&rdquo; action</h3>
    <p>Inside the new flow, click <strong>+ New step</strong>, then search for and select:</p>
    <p><code>Office 365 Outlook</code> &rarr; <code>Set up automatic replies (V2)</code></p>
    <p>If prompted to sign in, use <code>");
        sb.Append(userEsc);
        sb.Append(@"</code>. Then fill the action with the values below.</p>
  </div>

  <div class=""step""><span class=""num"">4</span>
    <h3>Fill in the action fields</h3>
    <p>Tap the <strong>fx</strong> tab inside any field that takes an expression, paste the expression there, and click <strong>OK</strong>.</p>

    <h4>Status</h4>
    <p>Set to <code>Scheduled</code>.</p>

    <h4>Time zone</h4>
    <p>Set to <code>");
        sb.Append(Escape(tzId));
        sb.Append(@"</code> (");
        sb.Append(Escape(tzDisplay));
        sb.Append(@"). The expressions below produce times in this zone, so this match is what makes Outlook show the right wall-clock window.</p>

    <h4>Start time</h4>
    <div class=""copy-row"">
      <pre>");
        sb.Append(Escape(startTimeExpression));
        sb.Append(@"</pre>
      <button class=""copy"" data-target=""start"">Copy</button>
    </div>
    <textarea id=""start"" style=""display:none"">");
        sb.Append(Escape(startTimeExpression));
        sb.Append(@"</textarea>

    <h4>End time</h4>
    <div class=""copy-row"">
      <pre>");
        sb.Append(Escape(endTimeExpression));
        sb.Append(@"</pre>
      <button class=""copy"" data-target=""end"">Copy</button>
    </div>
    <textarea id=""end"" style=""display:none"">");
        sb.Append(Escape(endTimeExpression));
        sb.Append(@"</textarea>

    <h4>External audience</h4>
    <p>Set to <code>");
        sb.Append(externalAudienceAll ? "All" : "Known");
        sb.Append(@"</code>.</p>

    <h4>Internal reply message</h4>
    <div class=""copy-row"">
      <pre id=""internal-pre"">");
        sb.Append(Escape(internalReply));
        sb.Append(@"</pre>
      <button class=""copy"" data-target=""internal"">Copy</button>
    </div>
    <textarea id=""internal"" style=""display:none"">");
        sb.Append(Escape(internalReply));
        sb.Append(@"</textarea>

    <h4>External reply message</h4>
    <div class=""copy-row"">
      <pre id=""external-pre"">");
        sb.Append(Escape(externalReply));
        sb.Append(@"</pre>
      <button class=""copy"" data-target=""external"">Copy</button>
    </div>
    <textarea id=""external"" style=""display:none"">");
        sb.Append(Escape(externalReply));
        sb.Append(@"</textarea>
  </div>

  <div class=""step""><span class=""num"">5</span>
    <h3>Save and test</h3>
    <p>Click <strong>Save</strong> at the top of the flow editor. Then click <strong>Test</strong> &rarr; <strong>Manually</strong> &rarr; <strong>Test</strong> to run it once and verify Outlook accepts the settings. You should see a green check on every step.</p>
    <div class=""callout success""><strong>Success criteria:</strong> Open Outlook on the web and check Settings &rarr; Mail &rarr; Automatic replies. The schedule shown there should match your expected next off-hours window.</div>
  </div>

  <h2>How it works</h2>
  <div class=""card"">
    <p>Every morning your flow runs at <code>");
        sb.Append(triggerHour.ToString("00") + ":" + triggerMinute.ToString("00"));
        sb.Append(@"</code> and tells Outlook: &ldquo;Schedule an OOF reply starting now and ending at the next configured start-of-work.&rdquo; Outlook then handles the actual on/off transitions server-side. Because the flow lives in Microsoft 365, none of your computers need to be on for this to keep working.</p>
    <p>You can keep using OofManager as before. The local app and the cloud flow update the same Outlook auto-reply settings, so whoever runs last wins &mdash; and they all compute the same window from the same schedule, so the result is consistent.</p>
  </div>

  <h2>Customize per day (optional)</h2>
  <div class=""card"">
    <p>If your work hours vary significantly across days, you can replace the End Time expression with a per-day version. The expression below is what OofManager generated for your current schedule &mdash; the numbers in the <code>switch</code> are how many days from each weekday until the next configured workday.</p>
    <div class=""copy-row"">
      <pre>");
        sb.Append(Escape(endTimeExpression));
        sb.Append(@"</pre>
      <button class=""copy"" data-target=""custom"">Copy</button>
    </div>
    <textarea id=""custom"" style=""display:none"">");
        sb.Append(Escape(endTimeExpression));
        sb.Append(@"</textarea>
    <p style=""margin-top:12px""><em>To regenerate this expression after editing your schedule in OofManager, click <strong>Generate cloud sync setup guide</strong> again in the app.</em></p>
  </div>

  <div class=""footer"">Generated ");
        sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
        sb.Append(@" by OofManager. This file is regenerated every time you click the button &mdash; it's safe to delete after setup.</div>
</div>

<script>
  document.querySelectorAll('button.copy').forEach(function(btn){
    btn.addEventListener('click', function(){
      var ta = document.getElementById(btn.dataset.target);
      if (!ta) return;
      navigator.clipboard.writeText(ta.value).then(function(){
        var orig = btn.textContent;
        btn.textContent = 'Copied!';
        btn.classList.add('copied');
        setTimeout(function(){ btn.textContent = orig; btn.classList.remove('copied'); }, 1400);
      });
    });
  });
</script>
</body>
</html>");

        return sb.ToString();
    }

    private static TimeSpan ComputeRepresentativeEnd(WorkScheduleSnapshot s)
    {
        // Earliest end across configured workdays. Falls back to 17:30 if
        // none are configured (so the guide still produces sensible output).
        TimeSpan? earliest = null;
        foreach (var d in WeekDays)
        {
            if (!s.IsWorkday(d)) continue;
            var e = s.GetEnd(d);
            if (earliest == null || e < earliest.Value) earliest = e;
        }
        return earliest ?? new TimeSpan(17, 30, 0);
    }

    private static TimeSpan ComputeRepresentativeStart(WorkScheduleSnapshot s)
    {
        TimeSpan? latest = null;
        foreach (var d in WeekDays)
        {
            if (!s.IsWorkday(d)) continue;
            var v = s.GetStart(d);
            if (latest == null || v > latest.Value) latest = v;
        }
        return latest ?? new TimeSpan(9, 0, 0);
    }

    /// <summary>
    /// Power Automate's <c>dayOfWeek</c> returns 0..6 with Sunday=0. We emit a
    /// nested <c>if</c> chain that maps each weekday to the right number of
    /// days to skip ahead to the next configured workday, then offsets to the
    /// work-start hour and serialises as ISO. The result is correct for the
    /// user's per-day workday config (e.g. Fri jumps 3 days to Monday for a
    /// classic Mon–Fri schedule).
    /// </summary>
    private static string BuildEndTimeExpression(int[] hopDays, int hour, int minute, string tzId)
    {
        // Build: if(equals(dow,0), 1, if(equals(dow,1), 1, if(... )))
        // where each number is hopDays[i] for that day-of-week index.
        var sb = new StringBuilder();
        for (int i = 0; i < 6; i++)
        {
            sb.Append("if(equals(dayOfWeek(utcNow()),").Append(i).Append("),").Append(hopDays[i]).Append(',');
        }
        sb.Append(hopDays[6]);
        for (int i = 0; i < 6; i++) sb.Append(')');

        var hopExpression = sb.ToString();

        // convertFromUtc(utcNow(), 'TZ') gives "now" in the user's local wall
        // clock; startOfDay anchors at local midnight. Add hopDays to land on
        // the next configured workday, then offset to the work-start hour.
        // The action's Time Zone field is set to the same TZ so Power Automate
        // sends the right absolute moment to Graph / Outlook.
        var localToday = $"startOfDay(convertFromUtc(utcNow(), '{tzId}'))";
        var full = $"formatDateTime(addMinutes(addHours(addDays({localToday}, {hopExpression}), {hour}), {minute}), 'yyyy-MM-ddTHH:mm:ss')";
        return full;
    }

    /// <summary>
    /// Start-time expression: "now" expressed in the user's local TZ so the
    /// action's Time Zone field interprets it as the same wall-clock instant.
    /// </summary>
    private static string BuildStartTimeExpression(string tzId)
    {
        return $"formatDateTime(convertFromUtc(utcNow(), '{tzId}'), 'yyyy-MM-ddTHH:mm:ss')";
    }

    private static string BuildWeeklyScheduleSummary(WorkScheduleSnapshot s)
    {
        // One inline span per day so the card stays compact.
        var sb = new StringBuilder();
        foreach (var d in WeekDays)
        {
            if (sb.Length > 0) sb.Append(" &middot; ");
            if (!s.IsWorkday(d))
            {
                sb.Append(@"<span style=""opacity:0.5"">").Append(d.ToString().Substring(0, 3)).Append(" off</span>");
            }
            else
            {
                sb.Append("<code>").Append(d.ToString().Substring(0, 3)).Append(' ')
                  .Append(s.GetStart(d).ToString(@"hh\:mm")).Append('-')
                  .Append(s.GetEnd(d).ToString(@"hh\:mm")).Append("</code>");
            }
        }
        return sb.ToString();
    }

    private static string BuildDaysSelectionLabel(WorkScheduleSnapshot s)
    {
        var days = WeekDays.Where(s.IsWorkday).Select(d => $"<code>{d}</code>").ToList();
        return days.Count == 0
            ? "(none — please enable at least one workday in OofManager)"
            : string.Join(", ", days);
    }

    private static readonly DayOfWeek[] WeekDays =
    {
        DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
        DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday,
    };

    private static bool HasPerDayVariation(WorkScheduleSnapshot s)
    {
        TimeSpan? start = null, end = null;
        foreach (var d in WeekDays)
        {
            if (!s.IsWorkday(d)) continue;
            if (start == null) { start = s.GetStart(d); end = s.GetEnd(d); continue; }
            if (s.GetStart(d) != start.Value || s.GetEnd(d) != end!.Value) return true;
        }
        return false;
    }

    private static string Escape(string s) => string.IsNullOrEmpty(s) ? string.Empty : WebUtility.HtmlEncode(s);
}
