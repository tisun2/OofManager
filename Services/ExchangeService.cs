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
    private readonly SemaphoreSlim _lock = new(1, 1);

    private bool _isConnected;
    private string? _userPrincipalName;

    public bool IsConnected => _isConnected;

    private Task? _prewarmTask;
    private bool _moduleImported;

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
            await InvokeAsync("Import-Module ExchangeOnlineManagement -DisableNameChecking -ErrorAction Stop");
            _moduleImported = true;
        }
        catch
        {
            // Pre-warm is best-effort. If it fails, ConnectAsync will surface the error.
            _moduleImported = false;
        }
    }

    public async Task ConnectAsync(string? upnHint = null)
    {
        // Wait for any in-flight pre-warm; if none was started, do the work now.
        if (_prewarmTask != null)
        {
            try { await _prewarmTask; } catch { _prewarmTask = null; }
        }
        EnsureRunspace();
        if (!_moduleImported)
        {
            await InvokeAsync("Import-Module ExchangeOnlineManagement -DisableNameChecking -ErrorAction Stop");
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
        var script = $@"
Connect-ExchangeOnline `
{upnLine}    -ShowBanner:$false `
    -SkipLoadingFormatData `
    -SkipLoadingCmdletHelp `
    -CommandName 'Get-MailboxAutoReplyConfiguration','Set-MailboxAutoReplyConfiguration' `
    -ErrorAction Stop | Out-Null
(Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
";

        var output = await InvokeAsync(script);

        _userPrincipalName = output
            .Select(o => o?.ToString()?.Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .LastOrDefault();

        if (string.IsNullOrWhiteSpace(_userPrincipalName))
        {
            // Connect-ExchangeOnline succeeded but Get-ConnectionInformation returned
            // nothing usable. Clean up the live REST session so a retry doesn't pile up
            // ghost sessions on the server side.
            try { await InvokeAsync("Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue"); }
            catch { /* best-effort */ }
            throw new Exception("Connected, but failed to read the current signed-in user.");
        }

        _isConnected = true;
    }

    public async Task DisconnectAsync()
    {
        try
        {
            if (_runspace != null && _runspace.RunspaceStateInfo.State == RunspaceState.Opened)
            {
                await InvokeAsync("Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue");
            }
        }
        catch { /* best-effort */ }
        finally
        {
            _isConnected = false;
            _userPrincipalName = null;
            _moduleImported = false;
            _prewarmTask = null;
            CloseRunspace();
        }
    }

    public Task<string> GetCurrentUserAsync()
    {
        return Task.FromResult(_userPrincipalName ?? "Unknown");
    }

    public async Task<OofSettings> GetOofSettingsAsync()
    {
        EnsureConnected();

        var upn = EscapePowerShellSingleQuotedString(_userPrincipalName!);
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

        var output = await InvokeAsync(script);
        var json = ExtractJson(string.Join("\n", output.Select(o => o?.ToString() ?? "")));
        var doc = JsonDocument.Parse(json);
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

        var state = settings.Status switch
        {
            OofStatus.Enabled => "Enabled",
            OofStatus.Scheduled => "Scheduled",
            _ => "Disabled"
        };
        var audience = settings.ExternalAudienceAll ? "All" : "Known";
        var upn = EscapePowerShellSingleQuotedString(_userPrincipalName!);
        var internalMsgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(PlainTextToHtml(settings.InternalReply)));
        var externalMsgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(PlainTextToHtml(settings.ExternalReply)));

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

        var output = await InvokeAsync(sb.ToString());
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

        await _lock.WaitAsync();
        try
        {
            // Use the native APM API instead of Task.Run(() => ps.Invoke()):
            // BeginInvoke/EndInvoke lets PowerShell schedule the work itself and avoids
            // tying up a thread-pool worker for the entire (potentially multi-second)
            // cmdlet duration — e.g. Connect-ExchangeOnline normally blocks for ~2s.
            var ps = PowerShell.Create();
            try
            {
                ps.Runspace = _runspace;
                ps.AddScript(script);
                // BeginInvoke/EndInvoke surfaces results as PSDataCollection<PSObject>;
                // wrap in Collection<PSObject> to keep this method's signature stable
                // for callers (which already enumerate the result via LINQ).
                var asyncResults = await Task.Factory.FromAsync(ps.BeginInvoke(), ps.EndInvoke);
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
                ps.Dispose();
            }
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
