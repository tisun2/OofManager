using System.Collections.ObjectModel;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OofManager.Wpf.Models;

namespace OofManager.Wpf.Services;

/// <summary>
/// Talks to Exchange Online by hosting Windows PowerShell 5.1 *in-process* through
/// System.Management.Automation. There is no powershell.exe child process and no
/// console window. Connect-ExchangeOnline's WAM dialog uses our WPF window as its
/// parent (via GetForegroundWindow fallback), which the tenant's token-protection
/// policy requires.
/// </summary>
public class ExchangeService : IExchangeService, IAsyncDisposable
{
    private Runspace? _runspace;
    private PowerShell? _ps;
    private readonly SemaphoreSlim _lock = new(1, 1);

    private bool _isConnected;
    private string? _userPrincipalName;
    private string? _targetMailboxIdentity;

    public bool IsConnected => _isConnected;

    private Task? _prewarmTask;
    private bool _moduleImported;
    // Cached background ConnectAsync. Populated by TryAutoConnectAsync from
    // App.OnStartup so the LoginPage's auto-sign-in path can simply await the
    // *same* connect attempt instead of kicking off a duplicate one.
    private Task? _autoConnectTask;

    public Task PrewarmAsync()
    {
        // Race-safe lazy init: if a prewarm is already in flight (or finished),
        // return that exact same Task so concurrent callers all wait on the
        // single ongoing import rather than spinning up duplicate runspaces.
        var existing = Volatile.Read(ref _prewarmTask);
        if (existing != null) return existing;
        var t = Task.Run(PrewarmCoreAsync);
        var prev = Interlocked.CompareExchange(ref _prewarmTask, t, null);
        return prev ?? t;
    }

    private async Task PrewarmCoreAsync()
    {
        try
        {
            EnsureRunspace();
            // Import once now so the user's click only pays the auth round-trip.
            // -DisableNameChecking skips a slow per-cmdlet name validation pass.
            // The module ships in the app's Modules\ folder so no install step is needed.
            await InvokeAsync("Import-Module ExchangeOnlineManagement -DisableNameChecking -ErrorAction Stop").ConfigureAwait(false);
            _moduleImported = true;
        }
        catch
        {
            // Pre-warm is best-effort. If it fails, ConnectAsync will surface the error.
            _moduleImported = false;
        }
    }

    public Task TryAutoConnectAsync(string upnHint)
    {
        // Already connected: nothing to do. The caller will observe IsConnected==true
        // immediately when it awaits this completed task.
        if (_isConnected) return Task.CompletedTask;

        var existing = Volatile.Read(ref _autoConnectTask);
        if (existing != null) return existing;

        var t = Task.Run(async () =>
        {
            try
            {
                await ConnectAsync(upnHint).ConfigureAwait(false);
            }
            catch
            {
                // Silent attempt: swallow. LoginViewModel inspects IsConnected after
                // awaiting and falls back to manual sign-in when this happens.
            }
            finally
            {
                // Drop the cached task on every outcome so a subsequent manual
                // Sign In can launch a fresh attempt without being short-circuited
                // by a stale completed/failed task.
                Volatile.Write(ref _autoConnectTask, null);
            }
        });
        var prev = Interlocked.CompareExchange(ref _autoConnectTask, t, null);
        return prev ?? t;
    }

    public async Task ConnectAsync(string? upnHint = null)
    {
        // Wait for any in-flight pre-warm; if none was started, do the work now.
        if (_prewarmTask != null)
        {
            try { await _prewarmTask.ConfigureAwait(false); } catch { _prewarmTask = null; }
        }
        EnsureRunspace();
        if (!_moduleImported)
        {
            await InvokeAsync("Import-Module ExchangeOnlineManagement -DisableNameChecking -ErrorAction Stop").ConfigureAwait(false);
            _moduleImported = true;
        }

        // Allocate the hidden console MSAL/WAM requires — only now, just before
        // the WAM dialog pops. Doing it here (not at app startup) keeps the app
        // launch flash-free; any brief conhost flicker is hidden under the WAM
        // auth dialog that appears immediately after.
        Program.EnsureHiddenConsole();

        // -CommandName limits which Exchange cmdlets are proxied (huge speedup: from ~10s
        // down to ~2s) since we only ever call the two AutoReplyConfiguration cmdlets.
        // Get-ConnectionInformation runs in the same script invocation to avoid an extra round-trip.
        // -UserPrincipalName: when supplied, MSAL/WAM tries to silently reuse a
        // cached refresh token for that account (no UI). If the cache is empty
        // or the token is no longer valid (password change, CA policy, 90-day
        // inactivity expiry), WAM falls back to its interactive sign-in dialog.
        var upnLine = string.IsNullOrWhiteSpace(upnHint)
            ? string.Empty
            : $"    -UserPrincipalName '{EscapePowerShellSingleQuotedString(upnHint!)}' `\n";
        // Also resolve the mailbox's PrimarySmtpAddress via the V3 REST cmdlet
        // Get-EXOMailbox. We use that as the -Identity for all subsequent
        // OOF reads/writes so OofManager hits the same EXO routing anchor that
        // Outlook desktop uses — without this, signing in with an alias UPN
        // (tisun@microsoft.com) made the mailbox locator pick a different
        // backend replica than Outlook's primary-SMTP (Tianyue.Sun@…) read,
        // producing the visible "OofManager says Scheduled, Outlook says
        // Disabled" inconsistency. Best-effort: if Get-EXOMailbox fails (perm
        // denied, REST endpoint unreachable, etc.) we fall back to the hint
        // or the connection UPN, which is the old behaviour.
        var script = $@"
Connect-ExchangeOnline `
{upnLine}    -ShowBanner:$false `
    -SkipLoadingFormatData `
    -SkipLoadingCmdletHelp `
    -CommandName 'Get-MailboxAutoReplyConfiguration','Set-MailboxAutoReplyConfiguration' `
    -ErrorAction Stop | Out-Null
$connUpn = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
$primarySmtp = $null
try {{
    $primarySmtp = (Get-EXOMailbox -Identity $connUpn -Properties PrimarySmtpAddress -ErrorAction Stop).PrimarySmtpAddress
}} catch {{
    # Best-effort: if this user lacks Get-EXOMailbox access (e.g. delegated
    # scenarios) keep the original UPN and let the rest of the app proceed.
}}
[PSCustomObject]@{{ UPN = $connUpn; PrimarySmtp = $primarySmtp }} | ConvertTo-Json -Compress
";

        var output = await InvokeAsync(script).ConfigureAwait(false);

        var rawJson = output
            .Select(o => o?.ToString()?.Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .LastOrDefault();

        string? primarySmtp = null;
        if (!string.IsNullOrWhiteSpace(rawJson))
        {
            try
            {
                using var doc = JsonDocument.Parse(ExtractJson(rawJson!));
                var root = doc.RootElement;
                if (root.TryGetProperty("UPN", out var upnEl) && upnEl.ValueKind == JsonValueKind.String)
                    _userPrincipalName = upnEl.GetString();
                if (root.TryGetProperty("PrimarySmtp", out var smtpEl) && smtpEl.ValueKind == JsonValueKind.String)
                    primarySmtp = smtpEl.GetString();
            }
            catch
            {
                // Old shape (plain UPN string) — keep backward compatibility
                // in case the cmdlet output format ever surprises us.
                _userPrincipalName = rawJson;
            }
        }

        if (string.IsNullOrWhiteSpace(_userPrincipalName))
        {
            // Connect-ExchangeOnline succeeded but Get-ConnectionInformation returned
            // nothing usable. Clean up the live REST session so a retry doesn't pile up
            // ghost sessions on the server side.
            try { await InvokeAsync("Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue").ConfigureAwait(false); }
            catch { /* best-effort */ }
            throw new Exception("Connected, but failed to read the current signed-in user.");
        }

        // Anchor precedence for OOF reads/writes:
        //   1. PrimarySmtpAddress from Get-EXOMailbox (matches Outlook desktop).
        //   2. Explicit upnHint (user passed it via Switch Account).
        //   3. Connection UPN from Get-ConnectionInformation.
        if (!string.IsNullOrWhiteSpace(primarySmtp))
            _targetMailboxIdentity = primarySmtp!.Trim();
        else if (!string.IsNullOrWhiteSpace(upnHint))
            _targetMailboxIdentity = upnHint!.Trim();
        else
            _targetMailboxIdentity = _userPrincipalName;
        _isConnected = true;
    }

    public async Task DisconnectAsync()
    {
        try
        {
            if (_runspace != null && _runspace.RunspaceStateInfo.State == RunspaceState.Opened)
            {
                await InvokeAsync("Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue").ConfigureAwait(false);
            }
        }
        catch { /* best-effort */ }
        finally
        {
            _isConnected = false;
            _userPrincipalName = null;
            _targetMailboxIdentity = null;
            _moduleImported = false;
            _prewarmTask = null;
            CloseRunspace();
        }
    }

    public Task<string> GetCurrentUserAsync()
    {
        return Task.FromResult(_userPrincipalName ?? "Unknown");
    }

    public Task<string> GetCurrentMailboxIdentityAsync()
    {
        return Task.FromResult(_targetMailboxIdentity ?? _userPrincipalName ?? "Unknown");
    }

    public async Task<OofSettings> GetOofSettingsAsync()
    {
        EnsureConnected();

        var upn = EscapePowerShellSingleQuotedString(_targetMailboxIdentity ?? _userPrincipalName!);
        var script = $@"
function ConvertTo-Utf8Base64([string]$value) {{
    if ($null -eq $value) {{ return '' }}
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($value))
}}
$oof = Get-MailboxAutoReplyConfiguration -Identity '{upn}'
[PSCustomObject]@{{
    Status                = $oof.AutoReplyState.ToString()
    InternalMessageBase64 = ConvertTo-Utf8Base64 $oof.InternalMessage
    ExternalMessageBase64 = ConvertTo-Utf8Base64 $oof.ExternalMessage
    ExternalAudience      = $oof.ExternalAudience.ToString()
    StartTime             = if ($oof.StartTime) {{ $oof.StartTime.ToString('o') }} else {{ $null }}
    EndTime               = if ($oof.EndTime) {{ $oof.EndTime.ToString('o') }} else {{ $null }}
}} | ConvertTo-Json -Depth 2 -Compress
";

        var output = await InvokeAsync(script).ConfigureAwait(false);
        var json = ExtractJson(string.Join("\n", output.Select(o => o?.ToString() ?? "")));
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var statusStr = root.GetProperty("Status").GetString() ?? "Disabled";
        var status = statusStr switch
        {
            "Enabled" => OofStatus.Enabled,
            "Scheduled" => OofStatus.Scheduled,
            _ => OofStatus.Disabled
        };

        var settings = new OofSettings
        {
            Status = status,
            InternalReply = HtmlToPlainText(DecodeBase64String(root.GetProperty("InternalMessageBase64").GetString())),
            ExternalReply = HtmlToPlainText(DecodeBase64String(root.GetProperty("ExternalMessageBase64").GetString())),
            ExternalAudienceAll = root.GetProperty("ExternalAudience").GetString() != "Known",
        };

        var startStr = root.GetProperty("StartTime").GetString();
        if (!string.IsNullOrEmpty(startStr) && DateTimeOffset.TryParse(startStr, out var start))
            settings.StartTime = start;

        var endStr = root.GetProperty("EndTime").GetString();
        if (!string.IsNullOrEmpty(endStr) && DateTimeOffset.TryParse(endStr, out var end))
            settings.EndTime = end;

        return settings;
    }

    public async Task SetOofSettingsAsync(OofSettings settings)
    {
        EnsureConnected();

        var expectedState = settings.Status switch
        {
            OofStatus.Enabled => "Enabled",
            OofStatus.Scheduled => "Scheduled",
            _ => "Disabled"
        };

        // Two-stage verification:
        //   1. WriteOnce sends the Set, then reads back AutoReplyState *in the
        //      same script* (same PSSession anchor) — cheap, catches tenant
        //      policy rejection and silent no-ops.
        //   2. After a short delay we issue a fresh Get from a separate cmdlet
        //      invocation. The EXO mailbox locator may anchor that call to a
        //      different backend replica than the write hit, so this is what
        //      catches replica-lag drift (the classic "OofManager says X,
        //      Outlook says Y" symptom).
        //
        // We re-read with progressive backoff (300ms → 1s → 2s) because most
        // replicas converge in under a second; only the long-tail cases ever
        // need the longer waits, and forcing every save to sit through several
        // seconds of fixed delay is a noticeable UX regression on the happy
        // path. If the reads never converge we issue one more write — replica
        // lag is the common cause but a one-shot lost write is also possible —
        // before giving up.
        var readDelays = new[]
        {
            TimeSpan.FromMilliseconds(300),
            TimeSpan.FromSeconds(1),
            TimeSpan.FromSeconds(2),
        };
        const int maxWriteAttempts = 2;
        string? lastObservedState = null;
        DateTimeOffset? lastObservedStart = null;
        DateTimeOffset? lastObservedEnd = null;

        for (int writeAttempt = 1; writeAttempt <= maxWriteAttempts; writeAttempt++)
        {
            await WriteOnceAsync(settings, expectedState).ConfigureAwait(false);
            SyncLogger.Write(
                $"    WriteOnce attempt={writeAttempt} target=({_targetMailboxIdentity}) " +
                $"state={expectedState} window={settings.StartTime:yyyy-MM-ddTHH:mm}..{settings.EndTime:yyyy-MM-ddTHH:mm}");

            foreach (var delay in readDelays)
            {
                await Task.Delay(delay).ConfigureAwait(false);
                var verify = await GetOofSettingsAsync().ConfigureAwait(false);
                lastObservedState = verify.Status switch
                {
                    OofStatus.Enabled => "Enabled",
                    OofStatus.Scheduled => "Scheduled",
                    _ => "Disabled"
                };
                lastObservedStart = verify.StartTime;
                lastObservedEnd = verify.EndTime;

                var match = StateMatches(settings, verify);
                SyncLogger.Write(
                    $"      verify after {delay.TotalMilliseconds:0}ms -> {lastObservedState} " +
                    $"{lastObservedStart:yyyy-MM-ddTHH:mm}..{lastObservedEnd:yyyy-MM-ddTHH:mm} match={match}");
                if (match)
                    return;
            }
        }

        throw new Exception(
            $"Save succeeded but the mailbox state did not converge after {maxWriteAttempts} write attempts " +
            $"(expected '{expectedState}', last observed '{lastObservedState ?? "unknown"}'" +
            (settings.Status == OofStatus.Scheduled
                ? $", expected start {settings.StartTime:o} got {lastObservedStart:o}, expected end {settings.EndTime:o} got {lastObservedEnd:o}"
                : "") +
            "). This usually means Exchange replicated the change to one replica but the read still hits a stale one. " +
            "Wait a minute and refresh; if it persists, the tenant policy may be silently overriding the value.");
    }

    private static bool StateMatches(OofSettings expected, OofSettings actual)
    {
        if (expected.Status != actual.Status) return false;
        if (expected.ExternalAudienceAll != actual.ExternalAudienceAll) return false;

        // Only the Scheduled mode actually persists start/end as user-visible
        // values. Enabled/Disabled writes a sentinel window that the user
        // never sees, so don't include it in the equality check.
        if (expected.Status == OofStatus.Scheduled)
        {
            // Exchange occasionally rounds to the nearest minute; a 60-second
            // tolerance keeps the verifier from spinning on a non-issue.
            if (!TimesApproximatelyEqual(expected.StartTime, actual.StartTime)) return false;
            if (!TimesApproximatelyEqual(expected.EndTime, actual.EndTime)) return false;
        }
        return true;
    }

    private static bool TimesApproximatelyEqual(DateTimeOffset? a, DateTimeOffset? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        return Math.Abs((a.Value - b.Value).TotalSeconds) <= 60;
    }

    private async Task WriteOnceAsync(OofSettings settings, string expectedState)
    {
        var audience = settings.ExternalAudienceAll ? "All" : "Known";
        var upn = EscapePowerShellSingleQuotedString(_targetMailboxIdentity ?? _userPrincipalName!);
        var internalMsgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(PlainTextToHtml(settings.InternalReply)));
        var externalMsgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(PlainTextToHtml(settings.ExternalReply)));

        var state = expectedState;
        var sb = new StringBuilder();
        sb.AppendLine($"$internalMessage = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{internalMsgBase64}'))");
        sb.AppendLine($"$externalMessage = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{externalMsgBase64}'))");
        sb.AppendLine("$params = @{");
        sb.AppendLine($"    Identity = '{upn}'");
        sb.AppendLine($"    AutoReplyState = '{state}'");
        sb.AppendLine("    InternalMessage = $internalMessage");
        sb.AppendLine("    ExternalMessage = $externalMessage");
        sb.AppendLine($"    ExternalAudience = '{audience}'");

        if (settings.Status == OofStatus.Scheduled)
        {
            var startStr = settings.StartTime?.ToString("yyyy-MM-ddTHH:mm:ss") ?? DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            var endStr = settings.EndTime?.ToString("yyyy-MM-ddTHH:mm:ss") ?? DateTime.Now.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss");
            sb.AppendLine($"    StartTime = '{startStr}'");
            sb.AppendLine($"    EndTime = '{endStr}'");
            // Always clear DeclineEventsForScheduledOOF: the auto-decline
            // checkbox was removed from the UI, but earlier builds let users
            // turn it on — forcing $false here ensures the flag is cleared on
            // the next sync for those mailboxes. We deliberately don't touch
            // DeclineAllEventsForScheduledOOF (which would *cancel* already-
            // accepted meetings) since that was never wired up.
            sb.AppendLine("    DeclineEventsForScheduledOOF = $false");
        }
        else
        {
            // Caller asked for a non-scheduled state (Enabled / Disabled).
            // Exchange keeps StartTime / EndTime around from the last
            // Scheduled push if we don't overwrite them, which is what makes
            // Outlook keep displaying the stale "Send replies during this time
            // period 17:30 → 09:00" UI even though it's no longer enforced.
            // Collapse the window to a one-day stretch in the distant past so
            // the schedule reads as already expired and Outlook hides it.
            // (Exchange requires EndTime to be strictly greater than StartTime,
            // so we space the two sentinels a day apart.)
            var sentinelStart = new DateTime(2000, 1, 1, 0, 0, 0).ToString("yyyy-MM-ddTHH:mm:ss");
            var sentinelEnd = new DateTime(2000, 1, 2, 0, 0, 0).ToString("yyyy-MM-ddTHH:mm:ss");
            sb.AppendLine($"    StartTime = '{sentinelStart}'");
            sb.AppendLine($"    EndTime = '{sentinelEnd}'");
        }

        sb.AppendLine("}");
        // -ErrorAction Stop turns any non-terminating error (e.g. tenant policy
        // rejection, parameter binding issue in the EXO V3 proxy function) into
        // a terminating exception so InvokeAsync surfaces it. Warnings are
        // captured separately because the REST proxy occasionally downgrades
        // real failures (e.g. "operation not permitted by policy") to warnings.
        sb.AppendLine("$exoWarnings = @()");
        sb.AppendLine("Set-MailboxAutoReplyConfiguration @params -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable +exoWarnings | Out-Null");
        sb.AppendLine("if ($exoWarnings.Count -gt 0) { throw ('Set-MailboxAutoReplyConfiguration warning: ' + ($exoWarnings -join '; ')) }");
        // Read the value back from the server within the same session and emit
        // it as the only output. We compare it below so a silent no-op (cmdlet
        // returns success but tenant didn't persist the change) is caught
        // instead of being surfaced to the user as a fake "✅ saved" message.
        sb.AppendLine($"(Get-MailboxAutoReplyConfiguration -Identity '{upn}' -ErrorAction Stop).AutoReplyState.ToString()");

        var output = await InvokeAsync(sb.ToString()).ConfigureAwait(false);
        var actualState = output
            .Select(o => o?.ToString()?.Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .LastOrDefault();

        if (!string.Equals(actualState, state, StringComparison.OrdinalIgnoreCase))
        {
            throw new Exception(
                $"Save request was sent but the server-side state was not updated (expected '{state}', got '{actualState ?? "unknown"}'). " +
                "Possible causes: tenant policy blocks the change, the signed-in account lacks write permission, or the ExchangeOnlineManagement module version is incompatible with the REST endpoint.");
        }
    }

    private void EnsureConnected()
    {
        if (!_isConnected || _runspace == null || _runspace.RunspaceStateInfo.State != RunspaceState.Opened)
            throw new InvalidOperationException("Not connected to Exchange Online. Please sign in first.");
    }

    // ---------- In-process PowerShell runspace ----------

    private void EnsureRunspace()
    {
        if (_runspace != null && _runspace.RunspaceStateInfo.State == RunspaceState.Opened) return;

        // Prefer the ExchangeOnlineManagement module bundled in the app's Modules\
        // folder. That way the app works offline and doesn't depend on the user's
        // PSGallery install. We also include the user's standard module paths as
        // fallbacks so a user-installed copy still works.
        var bundledModules = Path.Combine(AppContext.BaseDirectory, "Modules");
        var userDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var modulePaths = new[]
        {
            bundledModules,
            Path.Combine(userDocs, "WindowsPowerShell", "Modules"),
            Path.Combine(userDocs, "PowerShell", "Modules"),
            Path.Combine(programFiles, "WindowsPowerShell", "Modules"),
            Path.Combine(Environment.SystemDirectory, "WindowsPowerShell", "v1.0", "Modules")
        };
        var existing = Environment.GetEnvironmentVariable("PSModulePath") ?? string.Empty;
        var combined = string.Join(";", modulePaths.Concat(new[] { existing }).Where(p => !string.IsNullOrEmpty(p)));
        Environment.SetEnvironmentVariable("PSModulePath", combined);

        var iss = InitialSessionState.CreateDefault();
        // Disable the authorization manager so script files (e.g. inside the bundled
        // ExchangeOnlineManagement module) load regardless of the system execution policy.
        iss.AuthorizationManager = null;
        iss.ThreadOptions = PSThreadOptions.UseCurrentThread;

        var rs = RunspaceFactory.CreateRunspace(iss);
        rs.Open();
        _runspace = rs;
    }

    private void CloseRunspace()
    {
        try { _ps?.Dispose(); } catch { }
        _ps = null;
        try { _runspace?.Close(); } catch { }
        try { _runspace?.Dispose(); } catch { }
        _runspace = null;
    }

    /// <summary>
    /// Invokes a PowerShell script against the persistent runspace. Throws on errors.
    /// </summary>
    private async Task<Collection<PSObject>> InvokeAsync(string script)
    {
        EnsureRunspace();

        await _lock.WaitAsync().ConfigureAwait(false);
        try
        {
            // Reuse a single PowerShell instance bound to the persistent runspace.
            // PowerShell.Create() allocates a per-call object graph (pipeline, streams,
            // command collection) that the host re-builds even when the underlying
            // runspace is hot — Commands.Clear() + Streams.ClearStreams() resets only
            // the bits that vary per invocation, which is materially cheaper on the
            // tight automation loop.
            var ps = _ps;
            if (ps == null || ps.Runspace != _runspace)
            {
                try { _ps?.Dispose(); } catch { }
                ps = PowerShell.Create();
                ps.Runspace = _runspace;
                _ps = ps;
            }
            else
            {
                ps.Commands.Clear();
                ps.Streams.ClearStreams();
            }

            ps.AddScript(script);
            // BeginInvoke/EndInvoke lets PowerShell schedule the work itself and avoids
            // tying up a thread-pool worker for the entire (potentially multi-second)
            // cmdlet duration — e.g. Connect-ExchangeOnline normally blocks for ~2s.
            var asyncResults = await Task.Factory.FromAsync(ps.BeginInvoke(), ps.EndInvoke).ConfigureAwait(false);
            if (ps.HadErrors && ps.Streams.Error.Count > 0)
            {
                var err = ps.Streams.Error[0];
                var msg = err.Exception?.Message ?? err.ToString();
                throw new Exception(msg);
            }
            return new Collection<PSObject>(asyncResults);
        }
        finally
        {
            _lock.Release();
        }
    }

    public ValueTask DisposeAsync()
    {
        CloseRunspace();
        return default;
    }

    // ---------- Helpers ----------

    // Compiled once per process; the LoadAsync code path runs each of these
    // up to 11 times (internal+external reply), so caching saves a measurable
    // chunk of CPU on every refresh.
    private static readonly RegexOptions CompiledFast = RegexOptions.Compiled | RegexOptions.CultureInvariant;
    private static readonly Regex RxComments = new(@"<!--.*?-->", CompiledFast | RegexOptions.Singleline);
    private static readonly Regex RxScriptStyle = new(@"<\s*(script|style)[^>]*>.*?<\s*/\s*\1\s*>", CompiledFast | RegexOptions.IgnoreCase | RegexOptions.Singleline);
    private static readonly Regex RxBr = new(@"<\s*br\s*/?\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxBlockClose = new(@"</\s*(p|div|li|tr|table|h[1-6])\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxCellClose = new(@"</\s*(td|th)\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxLiOpen = new(@"<\s*li[^>]*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxAnyTag = new(@"<[^>]+>", CompiledFast);
    private static readonly Regex RxBidi = new(@"[\u200E\u200F\u061C\u202A-\u202E\u2066-\u2069]", CompiledFast);
    private static readonly Regex RxZeroWidth = new(@"[\u200B-\u200D\uFEFF\u2060]", CompiledFast);
    private static readonly Regex RxSoftHyphen = new(@"[\u00AD\u2028\u2029]", CompiledFast);
    private static readonly Regex RxUnicodeSpaces = new(@"[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]", CompiledFast);
    private static readonly Regex RxRunOfSpaces = new(@"[ \f\v]+", CompiledFast);
    private static readonly Regex RxSpacesAroundNl = new(@" *\n *", CompiledFast);
    private static readonly Regex RxBlankLines = new(@"\n{3,}", CompiledFast);

    private static string EscapePowerShellSingleQuotedString(string value)
    {
        return value.Replace("'", "''");
    }

    private static string DecodeBase64String(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;
        return Encoding.UTF8.GetString(Convert.FromBase64String(value));
    }

    private static string ExtractJson(string raw)
    {
        var trimmed = raw.Trim();
        if (trimmed.StartsWith("{") || trimmed.StartsWith("[")) return trimmed;
        var idx = trimmed.IndexOfAny(new[] { '{', '[' });
        if (idx >= 0) return trimmed.Substring(idx);
        return trimmed;
    }

    private static string HtmlToPlainText(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var text = value;
        text = RxComments.Replace(text, string.Empty);
        text = RxScriptStyle.Replace(text, string.Empty);
        text = RxBr.Replace(text, "\n");
        text = RxBlockClose.Replace(text, "\n");
        text = RxCellClose.Replace(text, " ");
        text = RxLiOpen.Replace(text, "- ");
        text = RxAnyTag.Replace(text, string.Empty);
        text = WebUtility.HtmlDecode(text);

        text = RxBidi.Replace(text, string.Empty);
        text = RxZeroWidth.Replace(text, string.Empty);
        text = RxSoftHyphen.Replace(text, string.Empty);

        text = RxUnicodeSpaces.Replace(text, " ");
        text = text.Replace("\t", " ");
        text = text.Replace("\r\n", "\n").Replace("\r", "\n");
        text = RxRunOfSpaces.Replace(text, " ");
        text = RxSpacesAroundNl.Replace(text, "\n");
        text = RxBlankLines.Replace(text, "\n\n");
        return text.Trim();
    }

    private static string PlainTextToHtml(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var encoded = WebUtility.HtmlEncode(value.Trim());
        encoded = encoded.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "<br>");
        return $"<html><body>{encoded}</body></html>";
    }
}
