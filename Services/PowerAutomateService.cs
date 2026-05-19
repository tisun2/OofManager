using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OofManager.Wpf.Services;

/// <summary>
/// Toggles the user's "OofManager Cloud Schedule" flow in Power Automate by hosting
/// the Microsoft.PowerApps.PowerShell module in a hidden powershell.exe child
/// process. Child process isolation is intentional: the module ships its own
/// Microsoft.Identity.Client.dll and net48 has a single AppDomain, so loading
/// it alongside EXO's MSAL would race the CLR's first-wins binding. Stdio is
/// deliberately NOT redirected — ADAL's interactive sign-in WebUI requires a
/// real console/STA pump that goes away the moment we redirect, which causes
/// silent indefinite hangs on cache-miss. The result JSON is returned via a
/// temp file path passed in OOFMGR_PA_RESULTFILE.
/// </summary>
public sealed class PowerAutomateService : IPowerAutomateService
{
    // The CloudSchedulePackageGenerator stamps every solution flow with
    // displayName "OofManager Cloud Schedule ({alias})". Matching on this prefix
    // is robust to alias differences and to re-imports.
    private const string FlowDisplayNamePrefix = "OofManager Cloud Schedule";
    private const string ImportCacheUpnKey = "PowerAutomate.Import.Environment.Upn";
    private const string ImportCacheEnvironmentIdKey = "PowerAutomate.Import.Environment.Id";
    private const string ImportCacheEnvironmentDisplayNameKey = "PowerAutomate.Import.Environment.DisplayName";
    private const string ImportCacheInstanceUrlKey = "PowerAutomate.Import.Environment.InstanceUrl";
    private const string FlowCacheUpnKey = "PowerAutomate.Flow.Upn";
    private const string FlowCacheEnvironmentIdKey = "PowerAutomate.Flow.Environment.Id";
    private const string FlowCacheDisplayNameKey = "PowerAutomate.Flow.DisplayName";
    private const string FlowCacheNameKey = "PowerAutomate.Flow.Name";

    private readonly IPreferencesService _prefs;

    public PowerAutomateService(IPreferencesService prefs)
    {
        _prefs = prefs;
    }

    public Task<PowerAutomateStatusResult> GetOofManagerFlowStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default, IProgress<string>? progress = null)
        => RunStatusAsync(upnHint, displayNameHint, expectedFlowDisplayName, ct, progress);

    public Task<PowerAutomateResult> DisableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default)
        => RunOperationAsync("Disable-Flow", upnHint, displayNameHint, expectedFlowDisplayName, progress, ct);

    public Task<PowerAutomateResult> EnableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default)
        => RunOperationAsync("Enable-Flow", upnHint, displayNameHint, expectedFlowDisplayName, progress, ct);

    public async Task<CloudScheduleImportResult> ImportCloudScheduleSolutionAsync(
        string solutionZipPath,
        string solutionUniqueName,
        Guid workflowId,
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        if (!File.Exists(solutionZipPath))
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.OtherError,
                $"Solution zip not found at '{solutionZipPath}'.",
                null, null, null, null);
        }

        var cachedEnvironment = GetCachedImportEnvironment(upnHint);
        if (cachedEnvironment != null)
        {
            progress?.Report("Using saved Power Platform environment. Skipping environment search...");
            var cachedResult = await RunImportScriptAsync(
                solutionZipPath,
                solutionUniqueName,
                workflowId,
                upnHint,
                displayNameHint,
                forceOverwrite,
                cachedEnvironment,
                progress,
                ct).ConfigureAwait(false);

            if (cachedResult.Outcome == CloudScheduleImportOutcome.Success)
            {
                SaveImportEnvironmentCache(upnHint, cachedResult);
                return cachedResult;
            }

            if (!ShouldRetryWithEnvironmentLookup(cachedResult))
                return cachedResult;

            ClearImportEnvironmentCache();
            progress?.Report("Saved environment did not work. Refreshing environment list...");
        }

        var result = await RunImportScriptAsync(
            solutionZipPath,
            solutionUniqueName,
            workflowId,
            upnHint,
            displayNameHint,
            forceOverwrite,
            null,
            progress,
            ct).ConfigureAwait(false);

        if (result.Outcome == CloudScheduleImportOutcome.Success)
            SaveImportEnvironmentCache(upnHint, result);

        return result;
    }

    private async Task<CloudScheduleImportResult> RunImportScriptAsync(
        string solutionZipPath,
        string solutionUniqueName,
        Guid workflowId,
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        CachedImportEnvironment? cachedEnvironment,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        var envVars = new Dictionary<string, string>
        {
            ["OOFMGR_PA_UPN"]          = upnHint ?? string.Empty,
            ["OOFMGR_PA_DISPLAYNAME"]  = displayNameHint ?? string.Empty,
            ["OOFMGR_PA_ZIPPATH"]      = solutionZipPath,
            ["OOFMGR_PA_SOLNAME"]      = solutionUniqueName,
            ["OOFMGR_PA_WORKFLOWID"]   = workflowId.ToString("D"),
            ["OOFMGR_PA_FORCE"]        = forceOverwrite ? "1" : "0",
            ["OOFMGR_PA_CACHED_INSTANCEURL"] = cachedEnvironment?.InstanceUrl ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVID"]       = cachedEnvironment?.EnvironmentId ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVDISPLAY"]  = cachedEnvironment?.EnvironmentDisplayName ?? string.Empty,
        };

        // pac auth + import can take longer than a flow toggle, especially on
        // first run where pac has to create an auth profile. Keep the outer cap
        // comfortably above the CLI's own import wait.
        var json = await RunPowerShellChildAsync(
            BuildImportScript(),
            envVars,
            timeout: TimeSpan.FromMinutes(15),
            ct: ct,
            progress: progress).ConfigureAwait(false);

        return ParseImportResult(json.Json, json.ExitCode);
    }

    private CachedImportEnvironment? GetCachedImportEnvironment(string? upnHint)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null) return null;

        var cachedUpn = _prefs.GetString(ImportCacheUpnKey);
        if (!string.Equals(cachedUpn, normalizedUpn, StringComparison.OrdinalIgnoreCase))
            return null;

        var instanceUrl = _prefs.GetString(ImportCacheInstanceUrlKey);
        if (string.IsNullOrWhiteSpace(instanceUrl)) return null;

        instanceUrl = instanceUrl!.Trim().TrimEnd('/');
        if (!Uri.TryCreate(instanceUrl, UriKind.Absolute, out var uri) ||
            !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            ClearImportEnvironmentCache();
            return null;
        }

        return new CachedImportEnvironment(
            instanceUrl,
            _prefs.GetString(ImportCacheEnvironmentIdKey),
            _prefs.GetString(ImportCacheEnvironmentDisplayNameKey));
    }

    private void SaveImportEnvironmentCache(string? upnHint, CloudScheduleImportResult result)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(result.InstanceUrl)) return;

        using var batch = _prefs.BeginBatch();
        _prefs.Set(ImportCacheUpnKey, normalizedUpn);
        _prefs.Set(ImportCacheInstanceUrlKey, result.InstanceUrl!.Trim().TrimEnd('/'));
        _prefs.Set(ImportCacheEnvironmentIdKey, result.EnvironmentId);
        _prefs.Set(ImportCacheEnvironmentDisplayNameKey, result.EnvironmentDisplayName);
    }

    private void ClearImportEnvironmentCache()
    {
        using var batch = _prefs.BeginBatch();
        _prefs.Set(ImportCacheUpnKey, (string?)null);
        _prefs.Set(ImportCacheInstanceUrlKey, (string?)null);
        _prefs.Set(ImportCacheEnvironmentIdKey, (string?)null);
        _prefs.Set(ImportCacheEnvironmentDisplayNameKey, (string?)null);
        ClearFlowReferenceCache();
    }

    private CachedFlowReference? GetCachedFlowReference(string? upnHint, string expectedFlowDisplayName, string? expectedEnvironmentId)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(expectedFlowDisplayName)) return null;

        var cachedUpn = _prefs.GetString(FlowCacheUpnKey);
        if (!string.Equals(cachedUpn, normalizedUpn, StringComparison.OrdinalIgnoreCase))
            return null;

        var cachedDisplayName = _prefs.GetString(FlowCacheDisplayNameKey);
        if (!string.Equals(cachedDisplayName, expectedFlowDisplayName, StringComparison.OrdinalIgnoreCase))
            return null;

        var environmentId = _prefs.GetString(FlowCacheEnvironmentIdKey);
        if (string.IsNullOrWhiteSpace(environmentId)) return null;

        if (!string.IsNullOrWhiteSpace(expectedEnvironmentId) &&
            !string.Equals(environmentId, expectedEnvironmentId, StringComparison.OrdinalIgnoreCase))
            return null;

        var flowName = _prefs.GetString(FlowCacheNameKey);
        if (string.IsNullOrWhiteSpace(flowName)) return null;

        return new CachedFlowReference(environmentId!.Trim(), flowName!.Trim(), cachedDisplayName!.Trim());
    }

    private void SaveFlowReferenceCache(string? upnHint, string expectedFlowDisplayName, IReadOnlyList<PowerAutomateFlowReference> flowReferences)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(expectedFlowDisplayName)) return;

        var flowReference = flowReferences.FirstOrDefault(f =>
            !string.IsNullOrWhiteSpace(f.EnvironmentName) &&
            !string.IsNullOrWhiteSpace(f.FlowName) &&
            string.Equals(f.DisplayName, expectedFlowDisplayName, StringComparison.OrdinalIgnoreCase));
        if (flowReference is null) return;

        using var batch = _prefs.BeginBatch();
        _prefs.Set(FlowCacheUpnKey, normalizedUpn);
        _prefs.Set(FlowCacheEnvironmentIdKey, flowReference.EnvironmentName.Trim());
        _prefs.Set(FlowCacheDisplayNameKey, flowReference.DisplayName.Trim());
        _prefs.Set(FlowCacheNameKey, flowReference.FlowName.Trim());
    }

    private void ClearFlowReferenceCache()
    {
        _prefs.Set(FlowCacheUpnKey, (string?)null);
        _prefs.Set(FlowCacheEnvironmentIdKey, (string?)null);
        _prefs.Set(FlowCacheDisplayNameKey, (string?)null);
        _prefs.Set(FlowCacheNameKey, (string?)null);
    }

    private static string? NormalizeCacheUpn(string? upnHint)
    {
        return string.IsNullOrWhiteSpace(upnHint) ? null : upnHint!.Trim();
    }

    private static bool ShouldRetryWithEnvironmentLookup(CloudScheduleImportResult result)
    {
        if (result.Outcome == CloudScheduleImportOutcome.Success ||
            result.Outcome == CloudScheduleImportOutcome.AlreadyExists)
            return false;

        return result.Outcome != CloudScheduleImportOutcome.OtherError ||
            result.Message.IndexOf("Power Platform CLI", StringComparison.OrdinalIgnoreCase) < 0;
    }

    private sealed class CachedImportEnvironment
    {
        public CachedImportEnvironment(string instanceUrl, string? environmentId, string? environmentDisplayName)
        {
            InstanceUrl = instanceUrl;
            EnvironmentId = environmentId;
            EnvironmentDisplayName = environmentDisplayName;
        }

        public string InstanceUrl { get; }
        public string? EnvironmentId { get; }
        public string? EnvironmentDisplayName { get; }
    }

    private sealed class CachedFlowReference
    {
        public CachedFlowReference(string environmentName, string flowName, string displayName)
        {
            EnvironmentName = environmentName;
            FlowName = flowName;
            DisplayName = displayName;
        }

        public string EnvironmentName { get; }
        public string FlowName { get; }
        public string DisplayName { get; }
    }

    private async Task<PowerAutomateStatusResult> RunStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct, IProgress<string>? progress)
    {
        var json = await RunFlowScriptAsync("Get-Status", upnHint, displayNameHint, expectedFlowDisplayName, ct, progress).ConfigureAwait(false);
        var result = ParseStatusResult(json.Json, json.ExitCode);
        if (result.Outcome == PowerAutomateOutcome.Success)
            SaveFlowReferenceCache(upnHint, expectedFlowDisplayName, result.FlowReferences);
        else if (result.Outcome == PowerAutomateOutcome.NoFlowFound)
            ClearFlowReferenceCache();
        return result;
    }

    private async Task<PowerAutomateResult> RunOperationAsync(string verb, string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress, CancellationToken ct)
    {
        var json = await RunFlowScriptAsync(verb, upnHint, displayNameHint, expectedFlowDisplayName, ct, progress).ConfigureAwait(false);
        var result = ParseResult(json.Json, json.ExitCode);
        if (result.Outcome == PowerAutomateOutcome.Success)
            SaveFlowReferenceCache(upnHint, expectedFlowDisplayName, result.FlowReferences);
        else if (result.Outcome == PowerAutomateOutcome.NoFlowFound)
            ClearFlowReferenceCache();
        return result;
    }

    private Task<ChildResult> RunFlowScriptAsync(string verb, string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct, IProgress<string>? progress = null)
    {
        var cachedEnvironment = GetCachedImportEnvironment(upnHint);
        var cachedFlowReference = GetCachedFlowReference(upnHint, expectedFlowDisplayName, cachedEnvironment?.EnvironmentId);
        var envVars = new Dictionary<string, string>
        {
            ["OOFMGR_PA_VERB"]              = verb,
            ["OOFMGR_PA_PREFIX"]            = FlowDisplayNamePrefix,
            ["OOFMGR_PA_UPN"]               = upnHint ?? string.Empty,
            ["OOFMGR_PA_DISPLAYNAME"]       = displayNameHint ?? string.Empty,
            // Exact display name of the user's own OofManager Cloud Schedule flow,
            // e.g. "OofManager Cloud Schedule (TianyueSun)". Computed by
            // CloudSchedulePackageGenerator.ComputeFlowIdentity from the UPN's
            // sanitised local-part, so it's stable per user and never collides
            // with other "OofManager Cloud Schedule …" flows in the same env that
            // belong to colleagues or unrelated automations. We don't match by
            // workflow GUID because Power Automate re-assigns the workflowid
            // at solution-import time and doesn't preserve the one we stamped
            // into solution.xml.
            ["OOFMGR_PA_FLOWDISPLAYNAME"]   = expectedFlowDisplayName ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVID"]      = cachedEnvironment?.EnvironmentId ?? cachedFlowReference?.EnvironmentName ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVDISPLAY"] = cachedEnvironment?.EnvironmentDisplayName ?? string.Empty,
            ["OOFMGR_PA_CACHED_FLOWNAME"]   = cachedFlowReference?.FlowName ?? string.Empty,
        };
        return RunPowerShellChildAsync(
            BuildScript(),
            envVars,
            timeout: TimeSpan.FromMinutes(5),
            ct: ct,
            progress: progress);
    }

    private readonly struct ChildResult
    {
        public ChildResult(string? json, int exitCode) { Json = json; ExitCode = exitCode; }
        public string? Json { get; }
        public int ExitCode { get; }
    }

    /// <summary>
    /// Launches a hidden powershell.exe child (via cmd.exe so Windows
    /// allocates a console — see the class XML comment for why), passes the
    /// caller-supplied <paramref name="envVars"/> plus the module path +
    /// result-file path, waits up to <paramref name="timeout"/>, and returns
    /// the contents of the result JSON file. Result parsing is the caller's
    /// responsibility because Disable/Enable and Import use different
    /// payload shapes.
    /// </summary>
    private async Task<ChildResult> RunPowerShellChildAsync(
        string script,
        IDictionary<string, string> envVars,
        TimeSpan timeout,
        CancellationToken ct,
        IProgress<string>? progress = null)
    {
        var modulesDir = Path.Combine(AppContext.BaseDirectory, "Modules");
        var psFile = Path.Combine(Path.GetTempPath(), $"oofmgr-pa-{Guid.NewGuid():N}.ps1");
        var resultFile = Path.Combine(Path.GetTempPath(), $"oofmgr-pa-result-{Guid.NewGuid():N}.json");
        var progressFile = Path.Combine(Path.GetTempPath(), $"oofmgr-pa-progress-{Guid.NewGuid():N}.log");
        try
        {
            // UTF-8 WITH BOM is mandatory for Windows PowerShell 5.1 — without
            // a BOM it reads .ps1 files using the system ANSI code page, so
            // any non-ASCII char in the script (e.g. an emoji in a user-facing
            // error message) gets mojibake'd into an invalid byte sequence and
            // the script fails with a parse error and exit code 1 before any
            // line runs. The BOM is cheap, harmless, and future-proofs the
            // script against accidental Unicode additions.
            File.WriteAllText(psFile, script, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

            var psi = new ProcessStartInfo
            {
                // Launch via cmd.exe so Windows actually allocates a new
                // console for the powershell child. When a WPF (no-console)
                // app spawns powershell.exe directly with UseShellExecute=
                // false, the child sometimes ends up with no conhost at all
                // — and ADAL's interactive sign-in WPF dialog silently
                // refuses to render because it can't anchor an STA dispatcher
                // to a non-existent window station. cmd.exe is well-tested
                // for that hand-off: it always gets a console allocated, and
                // powershell inherits it.
                FileName = "cmd.exe",
                Arguments = $"/c powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"{psFile}\"",
                UseShellExecute = false,
                // Do NOT redirect stdio: ADAL's interactive WebUI launches
                // only when the host has a real console + STA pump.
                // Redirecting stdout/stderr makes the child headless and the
                // auth dialog never appears, leaving the runspace blocked.
                RedirectStandardOutput = false,
                RedirectStandardError = false,
                CreateNoWindow = false,
            };
            psi.EnvironmentVariables["OOFMGR_PA_MODULES"]    = modulesDir;
            psi.EnvironmentVariables["OOFMGR_PA_RESULTFILE"] = resultFile;
            psi.EnvironmentVariables["OOFMGR_PA_PROGRESSFILE"] = progressFile;
            foreach (var kv in envVars)
            {
                psi.EnvironmentVariables[kv.Key] = kv.Value;
            }

            using var proc = new Process { StartInfo = psi };
            proc.Start();
            var deadline = DateTime.UtcNow + timeout;
            var reportedProgressLines = 0;

            using (ct.Register(() => { try { if (!proc.HasExited) proc.Kill(); } catch { } }))
            {
                while (!proc.HasExited)
                {
                    ReportProgressLines(progressFile, ref reportedProgressLines, progress);
                    var remaining = deadline - DateTime.UtcNow;
                    if (remaining <= TimeSpan.Zero)
                    {
                        try { proc.Kill(); } catch { }
                        return new ChildResult(null, -1);
                    }

                    var delay = remaining < TimeSpan.FromMilliseconds(500) ? remaining : TimeSpan.FromMilliseconds(500);
                    await Task.Delay(delay, ct).ConfigureAwait(false);
                }
                proc.WaitForExit();
                ReportProgressLines(progressFile, ref reportedProgressLines, progress);
            }

            string? resultJson = null;
            try { if (File.Exists(resultFile)) resultJson = File.ReadAllText(resultFile, Encoding.UTF8); } catch { }
            return new ChildResult(resultJson, proc.ExitCode);
        }
        finally
        {
            try { File.Delete(psFile); } catch { /* best-effort */ }
            try { File.Delete(resultFile); } catch { /* best-effort */ }
            try { File.Delete(progressFile); } catch { /* best-effort */ }
        }
    }

    private static void ReportProgressLines(string progressFile, ref int reportedLineCount, IProgress<string>? progress)
    {
        if (progress is null || !File.Exists(progressFile)) return;

        string[] lines;
        try { lines = File.ReadAllLines(progressFile, Encoding.UTF8); }
        catch { return; }

        for (var i = reportedLineCount; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (!string.IsNullOrWhiteSpace(line)) progress.Report(line);
        }
        reportedLineCount = lines.Length;
    }

    private static CloudScheduleImportResult ParseImportResult(string? resultJson, int exitCode)
    {
        if (exitCode == -1 && resultJson is null)
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.SignInFailed,
                "Power Automate sign-in / pac import did not complete within 15 minutes. Falling back to the manual setup flow.",
                null, null, null, null);
        }
        if (string.IsNullOrWhiteSpace(resultJson))
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.OtherError,
                $"PowerShell child process exited with code {exitCode} but produced no result file.",
                null, null, null, null);
        }
        try
        {
            using var doc = JsonDocument.Parse(resultJson!);
            var root = doc.RootElement;
            string? Read(string n) => root.TryGetProperty(n, out var el) && el.ValueKind == JsonValueKind.String ? el.GetString() : null;

            var outcome = Read("Outcome") switch
            {
                "Success"             => CloudScheduleImportOutcome.Success,
                "SignInFailed"        => CloudScheduleImportOutcome.SignInFailed,
                "NoOwnedEnvironment"  => CloudScheduleImportOutcome.NoOwnedEnvironment,
                "AlreadyExists"       => CloudScheduleImportOutcome.AlreadyExists,
                "ImportFailed"        => CloudScheduleImportOutcome.ImportFailed,
                "TenantBlocked"       => CloudScheduleImportOutcome.TenantBlocked,
                _                     => CloudScheduleImportOutcome.OtherError,
            };
            return new CloudScheduleImportResult(
                outcome,
                Read("Message") ?? string.Empty,
                Read("EnvironmentId"),
                Read("EnvironmentDisplayName"),
                Read("InstanceUrl"),
                Read("WorkflowId"));
        }
        catch (Exception ex)
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.OtherError,
                $"Failed to parse PowerShell import result: {ex.Message}. Raw: {resultJson}",
                null, null, null, null);
        }
    }

    private static PowerAutomateResult ParseResult(string? resultJson, int exitCode)
    {
        if (exitCode == -1 && resultJson is null)
        {
            return new PowerAutomateResult(
                PowerAutomateOutcome.SignInFailed,
                "Power Automate sign-in did not complete within 5 minutes. Falling back to the browser.",
                Array.Empty<string>());
        }
        if (string.IsNullOrWhiteSpace(resultJson))
        {
            return new PowerAutomateResult(
                PowerAutomateOutcome.OtherError,
                $"PowerShell child process exited with code {exitCode} but produced no result file. The Power Automate modules may have failed to load.",
                Array.Empty<string>());
        }

        try
        {
            using var doc = JsonDocument.Parse(resultJson!);
            var root = doc.RootElement;
            var outcomeStr = root.TryGetProperty("Outcome", out var oEl) ? oEl.GetString() : null;
            var message = root.TryGetProperty("Message", out var mEl) ? mEl.GetString() ?? "" : "";
            var flows = root.TryGetProperty("Flows", out var fEl) && fEl.ValueKind == JsonValueKind.Array
                ? fEl.EnumerateArray().Select(e => e.GetString() ?? "").Where(s => s.Length > 0).ToArray()
                : Array.Empty<string>();
            var flowReferences = ReadFlowReferences(root);

            var outcome = outcomeStr switch
            {
                "Success" => PowerAutomateOutcome.Success,
                "NoFlowFound" => PowerAutomateOutcome.NoFlowFound,
                "SignInFailed" => PowerAutomateOutcome.SignInFailed,
                "SolutionAwareBlocked" => PowerAutomateOutcome.SolutionAwareBlocked,
                _ => PowerAutomateOutcome.OtherError,
            };
            return new PowerAutomateResult(outcome, message, flows, flowReferences);
        }
        catch (Exception ex)
        {
            return new PowerAutomateResult(
                PowerAutomateOutcome.OtherError,
                $"Failed to parse PowerShell result: {ex.Message}. Raw: {resultJson}",
                Array.Empty<string>());
        }
    }

    private static PowerAutomateStatusResult ParseStatusResult(string? resultJson, int exitCode)
    {
        if (exitCode == -1 && resultJson is null)
        {
            return new PowerAutomateStatusResult(
                PowerAutomateOutcome.SignInFailed,
                PowerAutomateFlowState.Unknown,
                "Power Automate sign-in did not complete within 5 minutes.",
                Array.Empty<string>());
        }
        if (string.IsNullOrWhiteSpace(resultJson))
        {
            return new PowerAutomateStatusResult(
                PowerAutomateOutcome.OtherError,
                PowerAutomateFlowState.Unknown,
                $"PowerShell child process exited with code {exitCode} but produced no result file. The Power Automate modules may have failed to load.",
                Array.Empty<string>());
        }

        try
        {
            using var doc = JsonDocument.Parse(resultJson!);
            var root = doc.RootElement;
            var outcomeStr = root.TryGetProperty("Outcome", out var oEl) ? oEl.GetString() : null;
            var message = root.TryGetProperty("Message", out var mEl) ? mEl.GetString() ?? "" : "";
            var stateStr = root.TryGetProperty("State", out var sEl) ? sEl.GetString() : null;
            var flows = root.TryGetProperty("Flows", out var fEl) && fEl.ValueKind == JsonValueKind.Array
                ? fEl.EnumerateArray().Select(e => e.GetString() ?? "").Where(s => s.Length > 0).ToArray()
                : Array.Empty<string>();
            var flowReferences = ReadFlowReferences(root);

            var outcome = outcomeStr switch
            {
                "Success" => PowerAutomateOutcome.Success,
                "NoFlowFound" => PowerAutomateOutcome.NoFlowFound,
                "SignInFailed" => PowerAutomateOutcome.SignInFailed,
                "SolutionAwareBlocked" => PowerAutomateOutcome.SolutionAwareBlocked,
                _ => PowerAutomateOutcome.OtherError,
            };
            var state = stateStr switch
            {
                "On" => PowerAutomateFlowState.On,
                "Off" => PowerAutomateFlowState.Off,
                "NotFound" => PowerAutomateFlowState.NotFound,
                _ when outcome == PowerAutomateOutcome.NoFlowFound => PowerAutomateFlowState.NotFound,
                _ => PowerAutomateFlowState.Unknown,
            };
            return new PowerAutomateStatusResult(outcome, state, message, flows, flowReferences);
        }
        catch (Exception ex)
        {
            return new PowerAutomateStatusResult(
                PowerAutomateOutcome.OtherError,
                PowerAutomateFlowState.Unknown,
                $"Failed to parse PowerShell status result: {ex.Message}. Raw: {resultJson}",
                Array.Empty<string>());
        }
    }

    private static IReadOnlyList<PowerAutomateFlowReference> ReadFlowReferences(JsonElement root)
    {
        if (!root.TryGetProperty("FlowReferences", out var refsEl) || refsEl.ValueKind != JsonValueKind.Array)
            return Array.Empty<PowerAutomateFlowReference>();

        var refs = new List<PowerAutomateFlowReference>();
        foreach (var refEl in refsEl.EnumerateArray())
        {
            if (refEl.ValueKind != JsonValueKind.Object) continue;
            var environmentName = refEl.TryGetProperty("EnvironmentName", out var envEl) && envEl.ValueKind == JsonValueKind.String ? envEl.GetString() : null;
            var flowName = refEl.TryGetProperty("FlowName", out var flowEl) && flowEl.ValueKind == JsonValueKind.String ? flowEl.GetString() : null;
            var displayName = refEl.TryGetProperty("DisplayName", out var displayEl) && displayEl.ValueKind == JsonValueKind.String ? displayEl.GetString() : null;
            if (!string.IsNullOrWhiteSpace(environmentName) &&
                !string.IsNullOrWhiteSpace(flowName) &&
                !string.IsNullOrWhiteSpace(displayName))
            {
                refs.Add(new PowerAutomateFlowReference(environmentName!, flowName!, displayName!));
            }
        }

        return refs;
    }

    /// <summary>
    /// The PowerShell payload that runs inside the child process. Reads inputs
    /// from env vars set by the caller and writes a single JSON object to the
    /// file at OOFMGR_PA_RESULTFILE that <see cref="ParseResult"/> consumes.
    /// We deliberately don't write to stdout — stdio isn't being read.
    /// </summary>
    private static string BuildScript() => @"
$ErrorActionPreference = 'Stop'
$verb            = $env:OOFMGR_PA_VERB
$prefix          = $env:OOFMGR_PA_PREFIX
$flowDisplayName = $env:OOFMGR_PA_FLOWDISPLAYNAME
$modulesDir      = $env:OOFMGR_PA_MODULES
$upn             = $env:OOFMGR_PA_UPN
$displayName     = $env:OOFMGR_PA_DISPLAYNAME
$cachedEnvId     = $env:OOFMGR_PA_CACHED_ENVID
$cachedEnvDisplay= $env:OOFMGR_PA_CACHED_ENVDISPLAY
$cachedFlowName  = $env:OOFMGR_PA_CACHED_FLOWNAME
$resultFile      = $env:OOFMGR_PA_RESULTFILE
$progressFile    = $env:OOFMGR_PA_PROGRESSFILE

# Per-step debug trace so we can diagnose hangs. Each line is timestamped
# and flushed immediately (Add-Content closes the handle), so even if the
# script never returns we know exactly which step blocked.
$debugLog = Join-Path $env:TEMP 'oofmgr-pa-debug.log'
try { Remove-Item -LiteralPath $debugLog -ErrorAction SilentlyContinue } catch {}
function Trace($m) {
    try {
        $line = ('{0}  {1}' -f (Get-Date -Format o), $m)
        Add-Content -LiteralPath $debugLog -Value $line -Encoding UTF8
    } catch {}
}
function Report($m) {
    if (-not $m) { return }
    Trace (""progress: "" + $m)
    if ($progressFile) {
        try {
            [System.IO.File]::AppendAllText($progressFile, ([string]$m) + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
        } catch {}
    }
}
Trace ""start: pid=$PID verb=$verb prefix='$prefix' flowDisplayName='$flowDisplayName' upn='$upn' displayName='$displayName' cachedEnvId='$cachedEnvId' cachedEnvDisplay='$cachedEnvDisplay' cachedFlowName='$cachedFlowName' modulesDir='$modulesDir'""
$actionLabel = if ($verb -eq 'Disable-Flow') { 'Turning off' } elseif ($verb -eq 'Enable-Flow') { 'Turning on' } else { 'Checking' }
Report (""$actionLabel Power Automate flow. Preparing..."")

# Friendly banner so the visible console window has context. ADAL's sign-in
# WPF dialog won't render if the host hides its windows up front, so we
# accept a temporarily visible console as the price of working auth.
$Host.UI.RawUI.WindowTitle = 'OofManager - Power Automate sign-in'
Write-Host ''
$operationLabel = if ($verb -eq 'Get-Status') { 'check' } else { 'toggle' }
Write-Host ""  OofManager is signing you in to Power Automate to $operationLabel your Cloud Schedule flow."" -ForegroundColor Cyan
Write-Host '  If a sign-in dialog appears, complete it. This window closes automatically.' -ForegroundColor Cyan
Write-Host ''
Trace 'banner shown'

function New-FlowReference($flow) {
    return [ordered]@{
        EnvironmentName = [string]$flow.EnvironmentName
        FlowName        = [string]$flow.FlowName
        DisplayName     = [string]$flow.DisplayName
    }
}

function Get-FlowReferences($flows) {
    $refs = @()
    foreach ($flow in @($flows)) {
        if ($flow -and $flow.EnvironmentName -and $flow.FlowName -and $flow.DisplayName) {
            $refs += (New-FlowReference $flow)
        }
    }
    return @($refs)
}

function Emit($outcome, $message, $flows, $state = $null, $flowReferences = @()) {
    $obj = [ordered]@{
        Outcome = $outcome
        Message = $message
        Flows   = @($flows)
        State   = $state
        FlowReferences = @($flowReferences)
    }
    $json = ($obj | ConvertTo-Json -Compress)
    if ($resultFile) {
        [System.IO.File]::WriteAllText($resultFile, $json, [System.Text.UTF8Encoding]::new($false))
    }
    Trace ""emit: outcome=$outcome message=$message""
    exit 0
}

try {
    if ($modulesDir -and (Test-Path -LiteralPath $modulesDir)) {
        $env:PSModulePath = $modulesDir + ';' + $env:PSModulePath
        Trace ""PSModulePath prepended: $modulesDir""
    }

    $ProgressPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'

    # Only the user-scoped module is required: it provides Add-PowerAppsAccount,
    # Get-PowerAppEnvironment, Get-Flow, Disable-Flow, Enable-Flow. The
    # Administration module ships an incompatible Microsoft.Identity.Client.dll
    # which fails to load when the user module's copy is already in the
    # AppDomain — and we don't need any admin-only cmdlets here.
    Report 'Loading Power Automate tools...'
    Trace 'before Import-Module Microsoft.PowerApps.PowerShell'
    Import-Module Microsoft.PowerApps.PowerShell -DisableNameChecking -ErrorAction Stop | Out-Null
    Trace 'after  Import-Module Microsoft.PowerApps.PowerShell'

    # Sign in. With a valid cached refresh token (under
    # %LOCALAPPDATA%\Microsoft\PowerAppsCli) this returns silently; otherwise
    # ADAL pops its sign-in dialog as a separate top-level window.
    try {
        Report 'Signing in to Power Automate...'
        if ($upn) {
            Trace ""before Add-PowerAppsAccount -Username $upn""
            Add-PowerAppsAccount -Endpoint prod -Username $upn -ErrorAction Stop | Out-Null
        } else {
            Trace 'before Add-PowerAppsAccount (no upn)'
            Add-PowerAppsAccount -Endpoint prod -ErrorAction Stop | Out-Null
        }
        Trace 'after  Add-PowerAppsAccount'
        Report 'Signed in. Looking for your cloud schedule flow...'
    } catch {
        Trace ""Add-PowerAppsAccount FAILED: $($_.Exception.Message)""
        Emit 'SignInFailed' (""Sign-in to Power Automate failed: "" + $_.Exception.Message) @()
    }

    # Walk environments looking for flows whose displayName starts with our
    # prefix. We CANNOT enumerate every env the user can see — on a large
    # tenant (e.g. Microsoft corporate) Get-PowerAppEnvironment returns
    # hundreds of envs the user is merely visible-to, and Get-Flow on each
    # takes 1.5–3 s, so a naive foreach turns a 10 s click into a 25-minute
    # hang. We prioritise the most likely environments:
    #   1) Envs whose DisplayName contains the user's display name OR any
    #      token derived from the UPN (e.g. 'Tianyue Sun's Environment',
    #      'Sandy Sun's Environment' — most Microsoft employees have a
    #      personal dev env named after them, and that's where they import
    #      maker artefacts).
    #   2) The tenant's Default environment.
    #   3) At most $maxOtherEnvs additional envs, capped by total time.
    # The total budget needs to be generous enough that even on a 600+ env
    # tenant — where Get-PowerAppEnvironment alone runs ~45 s — there's
    # still time left to scan ~5-10 candidate envs after the listing.
    $maxOtherEnvs    = 30
    $totalBudgetSec  = 180
    $deadline        = (Get-Date).AddSeconds($totalBudgetSec)
    $matched         = @()

    function Add-MatchedFlow($envName, $flow, $sourceLabel) {
        if (-not $flow) { return $false }
        $dn = [string]$flow.DisplayName
        $flowName = [string]$flow.FlowName
        if ($dn -and $flowName -and $flowDisplayName -and ($dn -ieq $flowDisplayName)) {
            $script:matched += [PSCustomObject]@{
                EnvironmentName = $envName
                FlowName        = $flowName
                DisplayName     = $dn
                Enabled         = $flow.Enabled
            }
            Trace ""matched cloud flow by exact displayName via ${sourceLabel}: $dn flowName=$flowName enabled=$($flow.Enabled)""
            return $true
        }

        if ($dn -or $flowName) { Trace ""ignored flow from ${sourceLabel}: displayName='$dn' flowName='$flowName'"" }
        return $false
    }

    function Try-CachedFlow($envName, $envLabel, $flowName) {
        if (-not $envName -or -not $flowName) { return $false }
        try {
            Trace ""before Get-Flow cached $envLabel env=$envName flow=$flowName""
            $flow = Get-Flow -EnvironmentName $envName -FlowName $flowName -ErrorAction Stop
            Trace ""after  Get-Flow cached ${envLabel}""
            if (Add-MatchedFlow $envName $flow $envLabel) { return $true }
            Trace ""cached flow did not match expected displayName '$flowDisplayName'; falling back to environment scan""
        } catch {
            Trace ""cached Get-Flow failed ${envLabel}: $($_.Exception.Message)""
        }
        return $false
    }

    function Scan-Env($envName, $envLabel) {
        if ($matched.Count -gt 0) { return }
        if ((Get-Date) -gt $deadline) { return }
        try {
            Trace ""before Get-Flow $envLabel=$envName""
            $flows = @(Get-Flow -EnvironmentName $envName -ErrorAction SilentlyContinue)
            Trace ""after  Get-Flow ${envLabel}: count=$($flows.Count)""
        } catch { Trace ""Get-Flow EX: $($_.Exception.Message)""; $flows = @() }
        # Match on the EXACT display name (case-insensitive), e.g.
        # 'OofManager Cloud Schedule (TianyueSun)'. The suffix is derived from
        # the UPN's sanitised local-part by SanitizeAlias, so it's stable per
        # user and unique enough in practice that we won't accidentally toggle
        # someone else's flow that merely shares the 'OofManager Cloud Schedule'
        # prefix. We don't match on workflow GUID because Power Automate
        # doesn't preserve the workflowid we stamp into solution.xml at
        # import time — the runtime FlowName ends up being a different
        # GUID assigned by the solution-aware import pipeline.
        foreach ($f in $flows) { $null = Add-MatchedFlow $envName $f $envLabel }
    }

    function Convert-RawEnvironment($envObject) {
        if (-not $envObject) { return $null }
        return [PSCustomObject]@{
            EnvironmentName = [string]$envObject.name
            DisplayName     = [string]$envObject.properties.displayName
            IsDefault       = [bool]$envObject.properties.isDefault
            Location        = [string]$envObject.location
            CreatedTime     = $envObject.properties.createdTime
            CreatedBy       = [string]$envObject.properties.createdBy.userPrincipalName
            Internal        = $envObject
        }
    }

    function Test-DeveloperEnvironment($env) {
        if (-not $env -or -not $env.Internal -or -not $env.Internal.properties) { return $false }
        $sku  = [string]$env.Internal.properties.environmentSku
        $type = [string]$env.Internal.properties.environmentType
        return ($sku -ieq 'Developer' -or $type -ieq 'Developer')
    }

    function Test-AdminEnvironmentForCurrentUser($env) {
        if (-not $env -or -not $env.Internal -or -not $env.Internal.properties) { return $false }
        $perms = $env.Internal.properties.permissions
        if (-not $perms) { return $false }

        $names = @($perms.PSObject.Properties | ForEach-Object { [string]$_.Name })
        foreach ($adminPerm in @(
            'AdminEnvironment',
            'DeleteEnvironment',
            'ManageEnvironment',
            'WriteEnvironment',
            'UpdateEnvironment',
            'AssignEnvironmentRole',
            'ManageEnvironmentRoleAssignment',
            'CreateDatabase')) {
            if ($names -contains $adminPerm) { return $true }
        }
        return $false
    }

    function Get-ExactDisplayNameEnvs($envs) {
        if (-not $displayName -or ([string]$displayName).IndexOf('@') -ge 0) { return @() }
        $exactDisplayName = ([string]$displayName).Trim()
        if (-not $exactDisplayName) { return @() }
        return @($envs | Where-Object {
            $d = [string]$_.DisplayName
            $d -and ($d -ieq $exactDisplayName)
        })
    }

    function Get-DeveloperAdminEnvsFast() {
        $developerEnvs = @()
        try {
            # This is the first-pass narrow query: unlike Get-PowerAppEnvironment,
            # it asks the service for Developer envs before we ever enumerate the
            # whole tenant. Some tenants/API versions may reject this $filter; in
            # that case we trace and let the existing full-scan fallback handle it.
            $developerEnvUri = ""https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/environments?`$expand=permissions&`$filter=properties/environmentSku%20eq%20%27Developer%27&api-version={apiVersion}""
            Report 'Finding your Developer environments...'
            Trace 'before fast developer-env query'
            $developerResult = InvokeApi -Method GET -Route $developerEnvUri -ErrorAction Stop
            $developerEnvs = @($developerResult.value | ForEach-Object { Convert-RawEnvironment $_ } | Where-Object { $_ -and (Test-DeveloperEnvironment $_) })
            Trace ""after  fast developer-env query: count=$($developerEnvs.Count)""
        } catch {
            Trace ""fast developer-env query failed: $($_.Exception.Message)""
            return @()
        }

        $adminDeveloperEnvs = @($developerEnvs | Where-Object { Test-AdminEnvironmentForCurrentUser $_ })
        $exactDeveloperEnvs = @(Get-ExactDisplayNameEnvs $developerEnvs)
        Trace ""exact-display-name developer envs from fast query: count=$($exactDeveloperEnvs.Count)""
        Trace ""admin+developer envs from fast query: count=$($adminDeveloperEnvs.Count)""

        $orderedDeveloperEnvs = @()
        $seenDeveloperIds = @{}
        foreach ($e in $exactDeveloperEnvs) {
            $id = [string]$e.EnvironmentName
            if ($id -and -not $seenDeveloperIds.ContainsKey($id)) { $seenDeveloperIds[$id] = $true; $orderedDeveloperEnvs += $e }
        }
        foreach ($e in $adminDeveloperEnvs) {
            $id = [string]$e.EnvironmentName
            if ($id -and -not $seenDeveloperIds.ContainsKey($id)) { $seenDeveloperIds[$id] = $true; $orderedDeveloperEnvs += $e }
        }
        foreach ($e in $developerEnvs) {
            $id = [string]$e.EnvironmentName
            if ($id -and -not $seenDeveloperIds.ContainsKey($id)) { $seenDeveloperIds[$id] = $true; $orderedDeveloperEnvs += $e }
        }

        if ($orderedDeveloperEnvs.Count -gt 0) { return @($orderedDeveloperEnvs) }

        # Permission shape is not consistent across tenants. If the service did
        # return Developer envs but no obvious admin marker, still scan those
        # Developer envs first; this keeps the first pass narrow without risking
        # a false negative caused by a missing/renamed permission key.
        return @()
    }

    if ($cachedEnvId -and ([string]$cachedEnvId).Trim().Length -gt 0) {
        $cachedEnvId = ([string]$cachedEnvId).Trim()
        $cachedLabel = if ($cachedEnvDisplay -and ([string]$cachedEnvDisplay).Trim().Length -gt 0) { ([string]$cachedEnvDisplay).Trim() } else { $cachedEnvId }
        $cachedFlowName = if ($cachedFlowName -and ([string]$cachedFlowName).Trim().Length -gt 0) { ([string]$cachedFlowName).Trim() } else { $null }
        if ($cachedFlowName) {
            Report ""Checking saved cloud flow in '$cachedLabel'...""
            Trace ""cached flow candidate displayName='$cachedLabel' envId='$cachedEnvId' flowName='$cachedFlowName'""
            $null = Try-CachedFlow $cachedEnvId 'cached-flow' $cachedFlowName
        }

        if ($matched.Count -eq 0) {
            Report ""Checking saved environment '$cachedLabel'...""
            Trace ""cached env candidate displayName='$cachedLabel' id='$cachedEnvId'""
            Scan-Env $cachedEnvId 'cached-env'
        }
    }

    $fastDeveloperAdminEnvs = @()
    $fastDeveloperScanned = 0
    if ($matched.Count -eq 0) {
        $fastDeveloperAdminEnvs = @(Get-DeveloperAdminEnvsFast)
        if ($fastDeveloperAdminEnvs.Count -gt 0) { Report ""Checking $($fastDeveloperAdminEnvs.Count) Developer environment(s)..."" }
        foreach ($e in $fastDeveloperAdminEnvs) {
            if ($matched.Count -gt 0) { break }
            if ((Get-Date) -gt $deadline) { break }
            Report ""Checking Developer environment '$($e.DisplayName)'...""
            Trace ""fast developer/admin env[$fastDeveloperScanned] displayName='$($e.DisplayName)' id='$($e.EnvironmentName)'""
            Scan-Env $e.EnvironmentName ""fast-developer-admin-env[$fastDeveloperScanned]""
            $fastDeveloperScanned++
        }
    }
    if ($matched.Count -gt 0) {
        Trace 'matched before full environment scan; skipping full environment scan'
    }

    if ($matched.Count -eq 0) {

    # Build a list of name-tokens to match env DisplayNames against.
    # We can't rely on a single accurate ""display name"" hint because the
    # WPF view-model stores the UPN (e.g. ""Tianyue.Sun@microsoft.com"") in
    # the UserDisplayName slot. So derive likely friendly-name fragments
    # from both the explicit hint AND the UPN local-part:
    #   ""Tianyue.Sun@microsoft.com"" -> ""Tianyue.Sun"", ""Tianyue Sun"",
    #                                    ""Tianyue"", ""Sun""
    # An env named ""Tianyue Sun's Environment"" will hit on ""Sun"",
    # ""Tianyue"", and ""Tianyue Sun"".
    $nameTokens = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    function Add-NameToken($t) {
        if (-not $t) { return }
        $t = ([string]$t).Trim()
        if ($t.Length -lt 2) { return }
        $null = $nameTokens.Add($t)
    }
    function Expand-FromString($raw) {
        if (-not $raw) { return }
        $s = [string]$raw
        # Strip @domain so emails contribute just the local part.
        $at = $s.IndexOf('@')
        if ($at -ge 0) { $s = $s.Substring(0, $at) }
        Add-NameToken $s
        $spaced = $s -replace '[._\-]+', ' '
        Add-NameToken $spaced
        foreach ($p in ($spaced -split '\s+')) { Add-NameToken $p }
    }
    Expand-FromString $displayName
    Expand-FromString $upn
    Trace (""name tokens: "" + (@($nameTokens) -join '|'))

    # Get the full env list once and reuse — Get-PowerAppEnvironment is the
    # expensive call (multi-second), but we need DisplayName + IsDefault for
    # smart prioritisation so we can't avoid it.
    Report 'Refreshing visible Power Platform environments...'
    Trace 'before Get-PowerAppEnvironment (all)'
    $allEnvs = @(Get-PowerAppEnvironment -ErrorAction SilentlyContinue)
    Trace ""after  Get-PowerAppEnvironment: count=$($allEnvs.Count)""
    Report ""Found $($allEnvs.Count) environments. Checking likely matches...""

    # Diagnostic: dump the FIRST env's full property shape so we can see
    # what owner/principal fields are actually populated by this module
    # version. Schema varies between PowerApps PowerShell module versions
    # (some put owner email under Internal.properties.principal.email,
    # others Internal.properties.createdBy.userPrincipalName, etc.).
    if ($allEnvs.Count -gt 0) {
        try {
            $sample = $allEnvs[0]
            Trace ""=== SAMPLE ENV PROPERTIES (env 0) ===""
            $sample.PSObject.Properties | ForEach-Object {
                $val = $_.Value
                if ($val -is [string] -or $val -is [bool] -or $val -is [int]) {
                    Trace ""  $($_.Name) = $val""
                } elseif ($val -ne $null) {
                    Trace ""  $($_.Name) :: $($val.GetType().Name)""
                }
            }
            if ($sample.Internal) {
                try {
                    $json = $sample.Internal | ConvertTo-Json -Depth 6 -Compress
                    if ($json.Length -gt 4000) { $json = $json.Substring(0, 4000) + '...[truncated]' }
                    Trace ""  Internal-json: $json""
                } catch { Trace ""  Internal-json EX: $($_.Exception.Message)"" }
            }
            Trace '=== END SAMPLE ==='
        } catch { Trace ""sample dump EX: $($_.Exception.Message)"" }
    }

    # Helper: extract the env's owner-email and owner-displayName from
    # whichever schema variant this module version uses. Returns a
    # PSCustomObject with .Email and .DisplayName, both possibly $null.
    function Get-EnvOwner($env) {
        $email = $null; $name = $null
        try {
            $p = $env.Internal.properties
            if ($p) {
                # Most common Microsoft-tenant schema: principal owner.
                if ($p.principal -and $p.principal.email)     { $email = [string]$p.principal.email }
                if ($p.principal -and $p.principal.displayName){ $name  = [string]$p.principal.displayName }
                if (-not $email -and $p.principalOwner -and $p.principalOwner.email) { $email = [string]$p.principalOwner.email }
                if (-not $name  -and $p.principalOwner -and $p.principalOwner.displayName) { $name = [string]$p.principalOwner.displayName }
                if (-not $email -and $p.createdBy -and $p.createdBy.userPrincipalName)     { $email = [string]$p.createdBy.userPrincipalName }
                if (-not $name  -and $p.createdBy -and $p.createdBy.displayName)           { $name  = [string]$p.createdBy.displayName }
            }
        } catch {}
        return [PSCustomObject]@{ Email = $email; DisplayName = $name }
    }

    # 1a) Envs whose owner email matches our UPN — typically 1-3 envs the
    #     user actually created (personal dev env + maybe a shared one).
    #     This is dramatically better than name-matching on huge tenants.
    $ownedEnvs = @()
    if ($upn) {
        $ownedEnvs = @($allEnvs | Where-Object {
            $o = Get-EnvOwner $_
            $o.Email -and ($o.Email -eq $upn)
        })
        Trace ""owned-by-upn envs: count=$($ownedEnvs.Count) (upn=$upn)""
        foreach ($e in $ownedEnvs) {
            $o = Get-EnvOwner $e
            Trace ""  owned env: displayName='$($e.DisplayName)' owner='$($o.DisplayName)' <$($o.Email)>""
        }
    }

    $exactDisplayNameEnvs = @(Get-ExactDisplayNameEnvs $allEnvs)
    Trace ""exact-display-name envs: count=$($exactDisplayNameEnvs.Count)""

    # 1b) Build the candidate set: owner-matched envs + envs whose display
    #     name contains any derived name-token. This catches both 'real
    #     display name' envs (e.g. 'Sandy Sun's Environment' when owner
    #     email also matches) AND tenant-shared envs nominally named after
    #     the user but owned by an admin.
    $namedEnvs = @()
    if ($nameTokens.Count -gt 0) {
        # Also seed tokens from env-owner display names we found above,
        # so the friendly-name pass picks up 'Sandy Sun' even though the
        # UPN is 'tisun'.
        foreach ($e in $ownedEnvs) {
            $o = Get-EnvOwner $e
            if ($o.DisplayName) { Expand-FromString $o.DisplayName }
        }
        Trace (""name tokens (final): "" + (@($nameTokens) -join '|'))
        $namedEnvs = @($allEnvs | Where-Object {
            $d = [string]$_.DisplayName
            if (-not $d) { return $false }
            foreach ($t in $nameTokens) {
                if ($d.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
            }
            return $false
        })
        Trace ""name-matched envs: count=$($namedEnvs.Count)""
    }

    # Merge exact display-name + owned + token-name, dedupe, keep order.
    $priorityEnvs = @()
    $seenIds      = @{}
    foreach ($e in $exactDisplayNameEnvs) {
        $id = [string]$e.EnvironmentName
        if ($id -and -not $seenIds.ContainsKey($id)) { $seenIds[$id] = $true; $priorityEnvs += $e }
    }
    foreach ($e in $ownedEnvs) {
        $id = [string]$e.EnvironmentName
        if ($id -and -not $seenIds.ContainsKey($id)) { $seenIds[$id] = $true; $priorityEnvs += $e }
    }
    foreach ($e in $namedEnvs) {
        $id = [string]$e.EnvironmentName
        if ($id -and -not $seenIds.ContainsKey($id)) { $seenIds[$id] = $true; $priorityEnvs += $e }
    }
    Trace ""priority envs (exact+owned+named, deduped): count=$($priorityEnvs.Count)""
    $i = 0
    foreach ($e in $priorityEnvs) {
        if ($matched.Count -gt 0) { break }
        if ((Get-Date) -gt $deadline) { break }
        Report ""Checking likely environment '$($e.DisplayName)'...""
        Trace ""priority env[$i] displayName='$($e.DisplayName)' id='$($e.EnvironmentName)'""
        Scan-Env $e.EnvironmentName ""priority-env[$i]""
        $i++
    }
    $namedScanned = if ($namedEnvs) { $namedEnvs.Count } else { 0 }

    # 2) Default env — where solution packages most often land if the user
    #    didn't pick a specific maker env at import time.
    $defaultEnv = $null
    if ($matched.Count -eq 0 -and (Get-Date) -lt $deadline) {
        $defaultEnv = @($allEnvs | Where-Object { $_.IsDefault -eq $true }) | Select-Object -First 1
        if (-not $defaultEnv) {
            try { $defaultEnv = Get-PowerAppEnvironment -Default -ErrorAction SilentlyContinue } catch {}
        }
        if ($defaultEnv -and $defaultEnv.EnvironmentName) {
            Report 'Checking default environment...'
            Trace ""default env displayName='$($defaultEnv.DisplayName)' id='$($defaultEnv.EnvironmentName)'""
            Scan-Env $defaultEnv.EnvironmentName 'default-env'
        } else {
            Trace 'no default env found'
        }
    }
    $defaultScanned = if ($defaultEnv -and $defaultEnv.EnvironmentName) { 1 } else { 0 }

    # 3) Other envs, capped + time-budgeted. Skip ones already scanned.
    $otherScanned = 0
    if ($matched.Count -eq 0 -and (Get-Date) -lt $deadline) {
        $alreadyScanned = @($priorityEnvs | ForEach-Object { [string]$_.EnvironmentName })
        if ($defaultEnv) { $alreadyScanned += [string]$defaultEnv.EnvironmentName }
        $candidates = $allEnvs | Where-Object { $alreadyScanned -notcontains [string]$_.EnvironmentName } | Select-Object -First $maxOtherEnvs
        foreach ($e in $candidates) {
            if ($matched.Count -gt 0) { break }
            if ((Get-Date) -gt $deadline) { Trace 'time budget exhausted, stopping scan'; break }
            if (($otherScanned % 5) -eq 0) { Report ""Checking additional environments ($($otherScanned + 1) of up to $maxOtherEnvs)..."" }
            Scan-Env $e.EnvironmentName ""env[$otherScanned]""
            $otherScanned++
        }
        Trace ""scanned $otherScanned additional envs (cap=$maxOtherEnvs, budget=$totalBudgetSec s)""
    }

    if ($matched.Count -eq 0) {
        $ownedCount = if ($ownedEnvs) { $ownedEnvs.Count } else { 0 }
        $tokenList = if ($nameTokens.Count -gt 0) { ""'"" + ((@($nameTokens)) -join ""', '"") + ""'"" } else { ""(none)"" }
        $msg = ""No flow named '"" + $flowDisplayName + ""' was found.`r`n"" +
               ""Scanned: "" + $ownedCount + "" env(s) you own + "" + $namedScanned + "" env(s) matching "" + $tokenList +
               "", "" + $defaultScanned + "" default env, "" +
               $otherScanned + "" of "" + $allEnvs.Count + "" other env(s) (cap "" + $maxOtherEnvs + "", "" + $totalBudgetSec + ""s budget).`r`n"" +
               ""If you imported the solution into a different environment, open the browser and toggle it there. If you've never imported it, click 'Generate Solution Package' first.""
        Emit 'NoFlowFound' $msg @()
    }

    }

    function Read-CurrentEnabled([object]$flow) {
        $result = [pscustomobject]@{ Known = $false; Enabled = $null }
        try {
            $fresh = if ($flow.EnvironmentName) {
                Get-Flow -EnvironmentName $flow.EnvironmentName -FlowName $flow.FlowName -ErrorAction Stop
            } else {
                Get-Flow -FlowName $flow.FlowName -ErrorAction Stop
            }
            $freshEnabledProp = $fresh.PSObject.Properties['Enabled']
            $freshState = if ($fresh.PSObject.Properties['Internal'] -and $fresh.Internal -and $fresh.Internal.properties) { $fresh.Internal.properties.state } else { $null }
            if ($freshEnabledProp -ne $null -and $freshEnabledProp.Value -is [bool]) {
                $result.Known = $true
                $result.Enabled = [bool]$freshEnabledProp.Value
            }
            Trace ""fresh flow state: $($flow.DisplayName) enabled='$($result.Enabled)' known=$($result.Known) rawState='$freshState'""
        } catch {
            Trace ""fresh flow state read failed for $($flow.DisplayName): $($_.Exception.Message)""
        }
        return $result
    }

    if ($verb -eq 'Get-Status') {
        foreach ($f in $matched) {
            Report ""Found cloud schedule flow '$($f.DisplayName)'. Reading current state...""
            $freshState = Read-CurrentEnabled $f
            if ($freshState.Known) {
                $state = if ($freshState.Enabled) { 'On' } else { 'Off' }
                $lowerState = $state.ToLowerInvariant()
                Report ""Power Automate flow is $lowerState.""
                Emit 'Success' ('Power Automate flow is ' + $lowerState) @($f.DisplayName) $state (Get-FlowReferences @($f))
            }
        }

        Emit 'OtherError' 'Power Automate flow was found, but its current state could not be read.' @($matched | ForEach-Object { $_.DisplayName }) 'Unknown'
    }

    $changed = @()
    $alreadyDesired = @()
    $errors  = @()
    foreach ($f in $matched) {
        try {
            $desiredEnabled = ($verb -eq 'Enable-Flow')
            $targetState = if ($desiredEnabled) { 'on' } else { 'off' }
            $enabledProp = $f.PSObject.Properties['Enabled']
            if ($enabledProp -ne $null) {
                $enabledType = if ($enabledProp.Value -ne $null) { $enabledProp.Value.GetType().FullName } else { '<null>' }
                Trace ""reported flow state before ${verb}: $($f.DisplayName) enabledRaw='$($enabledProp.Value)' type='$enabledType' desiredEnabled=$desiredEnabled""
            }

            Report ""Checking whether '$($f.DisplayName)' is already $targetState...""
            $freshState = Read-CurrentEnabled $f
            if ($freshState.Known -and $freshState.Enabled -eq $desiredEnabled) {
                Trace ""skip $verb; fresh state already desired: $($f.DisplayName)""
                Report ""Power Automate flow is already $targetState. Finishing...""
                $alreadyDesired += $f.DisplayName
                continue
            }

            Report ""Turning '$($f.DisplayName)' $targetState...""
            Trace ""before $verb env=$($f.EnvironmentName) flow=$($f.FlowName)""
            if ($f.EnvironmentName) {
                & $verb -EnvironmentName $f.EnvironmentName -FlowName $f.FlowName -ErrorAction Stop | Out-Null
            } else {
                & $verb -FlowName $f.FlowName -ErrorAction Stop | Out-Null
            }
            Trace ""after  $verb OK: $($f.DisplayName)""
            Report 'Power Automate accepted the change. Finishing...'
            $changed += $f.DisplayName
        } catch {
            Trace ""$verb FAILED on $($f.DisplayName): $($_.Exception.Message)""
            $errors += ($f.DisplayName + "": "" + $_.Exception.Message)
        }
    }

    if ($changed.Count -eq 0 -and $alreadyDesired.Count -eq 0 -and $errors.Count -gt 0) {
        $msg = $errors -join '; '
        if ($msg -match 'solution' -or $msg -match 'managed' -or $msg -match 'not.*allowed') {
            Emit 'SolutionAwareBlocked' $msg @()
        }
        Emit 'OtherError' $msg @()
    }

    $action = if ($verb -eq 'Disable-Flow') { 'Turned off' } else { 'Turned on' }
    $state = if ($verb -eq 'Disable-Flow') { 'off' } else { 'on' }
    $msg = if ($changed.Count -gt 0) { ""$action Power Automate flow"" } else { ""Power Automate flow is already $state"" }
    if ($errors.Count -gt 0) { $msg = $msg + "" (with "" + $errors.Count + "" partial failure(s): "" + ($errors -join '; ') + "")"" }
    Report 'Done.'
    Emit 'Success' $msg @($changed + $alreadyDesired) $null (Get-FlowReferences $matched)

} catch {
    Trace ""OUTER CATCH: $($_.Exception.Message)""
    Emit 'OtherError' $_.Exception.Message @()
}
";

    /// <summary>
    /// Payload for <see cref="ImportCloudScheduleSolutionAsync"/>. Auths via the
    /// bundled Microsoft.PowerApps.PowerShell module the same way the toggle
    /// script does, picks the env whose owner/name matches the signed-in user,
    /// then lets the Power Platform CLI handle the Dataverse solution import
    /// with <c>pac solution import</c>. pac uses its existing WAM profile when
    /// available, which avoids the Dataverse token preauthorization block some
    /// tenants hit on the older direct Web API path.
    /// </summary>
    private static string BuildImportScript() => @"
$ErrorActionPreference = 'Stop'
$modulesDir   = $env:OOFMGR_PA_MODULES
$upn          = $env:OOFMGR_PA_UPN
$displayName  = $env:OOFMGR_PA_DISPLAYNAME
$resultFile   = $env:OOFMGR_PA_RESULTFILE
$progressFile = $env:OOFMGR_PA_PROGRESSFILE
$zipPath      = $env:OOFMGR_PA_ZIPPATH
$solName      = $env:OOFMGR_PA_SOLNAME
$workflowId   = $env:OOFMGR_PA_WORKFLOWID
$forceFlag    = $env:OOFMGR_PA_FORCE
$cachedInstanceUrl = $env:OOFMGR_PA_CACHED_INSTANCEURL
$cachedEnvId       = $env:OOFMGR_PA_CACHED_ENVID
$cachedEnvDisplay  = $env:OOFMGR_PA_CACHED_ENVDISPLAY

$debugLog = Join-Path $env:TEMP 'oofmgr-pa-debug.log'
try { Remove-Item -LiteralPath $debugLog -ErrorAction SilentlyContinue } catch {}
function Trace($m) {
    try {
        $line = ('{0}  {1}' -f (Get-Date -Format o), $m)
        Add-Content -LiteralPath $debugLog -Value $line -Encoding UTF8
    } catch {}
}
function Report($m) {
    if (-not $m) { return }
    Trace (""progress: "" + $m)
    if ($progressFile) {
        try {
            [System.IO.File]::AppendAllText($progressFile, ([string]$m) + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
        } catch {}
    }
}
Trace ""start: pid=$PID op=import upn='$upn' solName='$solName' workflowId='$workflowId' force='$forceFlag' zip='$zipPath'""
Report 'Preparing Power Automate import...'

$Host.UI.RawUI.WindowTitle = 'OofManager - Power Automate import'
Write-Host ''
Write-Host '  OofManager is importing your Cloud Schedule flow into Power Automate.' -ForegroundColor Cyan
Write-Host '  If a sign-in dialog appears, complete it. This window closes automatically.' -ForegroundColor Cyan
Write-Host ''

function Emit($outcome, $message, $envId, $envDisplay, $instanceUrl, $wfid) {
    $obj = [ordered]@{
        Outcome                = $outcome
        Message                = $message
        EnvironmentId          = $envId
        EnvironmentDisplayName = $envDisplay
        InstanceUrl            = $instanceUrl
        WorkflowId             = $wfid
    }
    $json = ($obj | ConvertTo-Json -Compress)
    if ($resultFile) {
        [System.IO.File]::WriteAllText($resultFile, $json, [System.Text.UTF8Encoding]::new($false))
    }
    Trace ""emit: outcome=$outcome message=$message envId=$envId instanceUrl=$instanceUrl wfid=$wfid""
    exit 0
}

try {
    $envId = $null
    $envDisplay = $null
    $instance = $null

    if ($cachedInstanceUrl -and ([string]$cachedInstanceUrl).Trim().Length -gt 0) {
        $instance = ([string]$cachedInstanceUrl).Trim().TrimEnd('/')
        $envId = if ($cachedEnvId) { [string]$cachedEnvId } else { $null }
        $envDisplay = if ($cachedEnvDisplay) { [string]$cachedEnvDisplay } else { $instance }
        Trace ""using cached Dataverse instance: id=$envId display='$envDisplay' instance=$instance""
        Report ""Using saved Power Platform environment '$envDisplay'. Preparing automatic import...""
    } else {
    if ($modulesDir -and (Test-Path -LiteralPath $modulesDir)) {
        $env:PSModulePath = $modulesDir + ';' + $env:PSModulePath
    }
    $ProgressPreference = 'SilentlyContinue'
    $WarningPreference  = 'SilentlyContinue'

    Import-Module Microsoft.PowerApps.PowerShell -DisableNameChecking -ErrorAction Stop | Out-Null
    Trace 'module loaded'
    Report 'Signing in to Power Automate...'

    try {
        if ($upn) { Add-PowerAppsAccount -Endpoint prod -Username $upn -ErrorAction Stop | Out-Null }
        else      { Add-PowerAppsAccount -Endpoint prod -ErrorAction Stop | Out-Null }
        Trace 'signed in'
        Report 'Signed in. Finding your Power Platform environment...'
    } catch {
        Trace ""sign-in FAILED: $($_.Exception.Message)""
        Emit 'SignInFailed' (""Sign-in to Power Automate failed: "" + $_.Exception.Message) $null $null $null $null
    }

    function Convert-RawEnvironment($envObject) {
        if (-not $envObject) { return $null }
        return [PSCustomObject]@{
            EnvironmentName = [string]$envObject.name
            DisplayName     = [string]$envObject.properties.displayName
            IsDefault       = [bool]$envObject.properties.isDefault
            Location        = [string]$envObject.location
            CreatedTime     = $envObject.properties.createdTime
            CreatedBy       = [string]$envObject.properties.createdBy.userPrincipalName
            Internal        = $envObject
        }
    }

    function Test-DeveloperEnvironment($candidateEnv) {
        $sku = $null; $type = $null
        try {
            $sku  = [string]$candidateEnv.Internal.properties.environmentSku
            $type = [string]$candidateEnv.Internal.properties.environmentType
        } catch {}
        return ($sku -ieq 'Developer' -or $type -ieq 'Developer')
    }

    function Test-AdminEnvironmentForCurrentUser($candidateEnv) {
        if (-not $candidateEnv -or -not $candidateEnv.Internal -or -not $candidateEnv.Internal.properties) { return $false }
        $perms = $candidateEnv.Internal.properties.permissions
        if (-not $perms) { return $false }

        $names = @($perms.PSObject.Properties | ForEach-Object { [string]$_.Name })
        foreach ($adminPerm in @(
            'AdminEnvironment',
            'DeleteEnvironment',
            'ManageEnvironment',
            'WriteEnvironment',
            'UpdateEnvironment',
            'AssignEnvironmentRole',
            'ManageEnvironmentRoleAssignment',
            'CreateDatabase')) {
            if ($names -contains $adminPerm) { return $true }
        }
        return $false
    }

    function Get-InstanceUrl($candidateEnv) {
        try {
            $iu = $candidateEnv.Internal.properties.linkedEnvironmentMetadata.instanceUrl
            if ($iu) { return ([string]$iu).TrimEnd('/') }
        } catch {}
        return $null
    }

    function Select-FastImportEnvironment($envs) {
        if (-not $envs -or -not $displayName -or ([string]$displayName).IndexOf('@') -ge 0) { return $null }
        $exactDisplayName = ([string]$displayName).Trim()
        if (-not $exactDisplayName) { return $null }

        $exactEnvs = @($envs | Where-Object {
            $d = [string]$_.DisplayName
            $d -and ($d -ieq $exactDisplayName)
        })
        Trace ""fast import exact-display-name developer envs: $($exactEnvs.Count)""

        foreach ($pool in @(
            @($exactEnvs | Where-Object { Test-AdminEnvironmentForCurrentUser $_ }),
            @($exactEnvs))) {
            foreach ($e in $pool) {
                $iu = Get-InstanceUrl $e
                if ($iu) { return $e }
            }
        }
        return $null
    }

    function Get-DeveloperEnvsFast() {
        try {
            $developerEnvUri = ""https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/environments?`$expand=permissions&`$filter=properties/environmentSku%20eq%20%27Developer%27&api-version={apiVersion}""
            Report 'Finding your Developer environments...'
            Trace 'before fast developer-env import query'
            $developerResult = InvokeApi -Method GET -Route $developerEnvUri -ErrorAction Stop
            $developerEnvs = @($developerResult.value | ForEach-Object { Convert-RawEnvironment $_ } | Where-Object { $_ -and (Test-DeveloperEnvironment $_) })
            Trace ""after  fast developer-env import query: count=$($developerEnvs.Count)""
            return @($developerEnvs)
        } catch {
            Trace ""fast developer-env import query failed: $($_.Exception.Message)""
            return @()
        }
    }

    $target = $null
    $fastDeveloperEnvs = @(Get-DeveloperEnvsFast)
    if ($fastDeveloperEnvs.Count -gt 0) {
        $target = Select-FastImportEnvironment $fastDeveloperEnvs
        if ($target) {
            Trace ""fast import target env: id=$($target.EnvironmentName) display='$($target.DisplayName)' instance=$(Get-InstanceUrl $target)""
            Report ""Selected '$($target.DisplayName)'. Preparing automatic import...""
        } else {
            Trace 'fast developer import query found no exact display-name Dataverse target; falling back to full environment list'
        }
    }

    if (-not $target) {
    Trace 'before Get-PowerAppEnvironment'
    $allEnvs = @(Get-PowerAppEnvironment -ErrorAction SilentlyContinue)
    Trace ""env count: $($allEnvs.Count)""
    Report ""Found $($allEnvs.Count) environments. Selecting your Dataverse environment...""

    function Get-EnvOwner($e) {
        $email = $null; $name = $null
        try {
            $p = $e.Internal.properties
            if ($p) {
                if ($p.principal -and $p.principal.email)       { $email = [string]$p.principal.email }
                if ($p.principal -and $p.principal.displayName) { $name  = [string]$p.principal.displayName }
                if (-not $email -and $p.principalOwner -and $p.principalOwner.email)       { $email = [string]$p.principalOwner.email }
                if (-not $name  -and $p.principalOwner -and $p.principalOwner.displayName) { $name  = [string]$p.principalOwner.displayName }
                if (-not $email -and $p.createdBy -and $p.createdBy.userPrincipalName)     { $email = [string]$p.createdBy.userPrincipalName }
                if (-not $name  -and $p.createdBy -and $p.createdBy.displayName)           { $name  = [string]$p.createdBy.displayName }
            }
        } catch {}
        return [PSCustomObject]@{ Email = $email; DisplayName = $name }
    }

    # Build name-tokens like the toggle does, so envs named 'Sandy Sun's
    # Environment' match even when the owner principal field is empty or

    function Test-DeveloperEnvironment($candidateEnv) {
        $sku = $null; $type = $null
        try {
            $sku  = [string]$candidateEnv.Internal.properties.environmentSku
            $type = [string]$candidateEnv.Internal.properties.environmentType
        } catch {}
        return ($sku -ieq 'Developer' -or $type -ieq 'Developer')
    }

    $developerEnvs = @($allEnvs | Where-Object { Test-DeveloperEnvironment $_ })
    Trace ""developer envs: $($developerEnvs.Count)""

    $matchEnvs = $allEnvs
    if ($developerEnvs.Count -gt 0) {
        $matchEnvs = $developerEnvs
        Report ""Found $($developerEnvs.Count) Developer environments. Matching your account display name...""
    } else {
        Trace 'no Developer environments found; matching across all visible environments'
    }
    # populated by an admin. Tokenises both the explicit display-name
    # hint (when WPF has the real friendly name) AND the UPN local-part.
    $nameTokens = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    function Add-NameToken($t) {
        if (-not $t) { return }
        $t = ([string]$t).Trim()
        if ($t.Length -lt 2) { return }
        $null = $nameTokens.Add($t)
    }
    function Expand-FromString($raw) {
        if (-not $raw) { return }
        $s = [string]$raw
        $at = $s.IndexOf('@')
        if ($at -ge 0) { $s = $s.Substring(0, $at) }
        Add-NameToken $s
        $spaced = $s -replace '[._\-]+', ' '
        Add-NameToken $spaced
        foreach ($p in ($spaced -split '\s+')) { Add-NameToken $p }
    }
    Expand-FromString $displayName
    Expand-FromString $upn

    # 1) Strict UPN match within the preferred environment pool.
    $ownedEnvs = @()
    if ($upn) {
        $ownedEnvs = @($matchEnvs | Where-Object {
            $o = Get-EnvOwner $_
            $o.Email -and ($o.Email -eq $upn)
        })
    }
    Trace ""owned-by-upn envs: $($ownedEnvs.Count)""

    # Seed extra tokens from owned-env owner display names so a UPN like
    # tianyue.sun also generates the token 'Sandy Sun' once we know it.
    foreach ($e in $ownedEnvs) {
        $o = Get-EnvOwner $e
        if ($o.DisplayName) { Expand-FromString $o.DisplayName }
    }
    Trace (""name tokens: "" + (@($nameTokens) -join '|'))

    # 2) Exact display-name match — personal developer environments are often
    # named after the account's friendly display name (e.g. 'Sandy Sun'). Prefer
    # this over loose token matching when WPF resolved the real display name.
    $exactDisplayNameEnvs = @()
    if ($displayName -and ([string]$displayName).IndexOf('@') -lt 0) {
        $exactDisplayName = ([string]$displayName).Trim()
        $exactDisplayNameEnvs = @($matchEnvs | Where-Object {
            $d = [string]$_.DisplayName
            $d -and ($d -eq $exactDisplayName)
        })
    }
    Trace ""exact-display-name envs: $($exactDisplayNameEnvs.Count)""

    # 3) Display-name token match — picks up 'Sandy Sun's Environment' style
    #    dev envs whose principal field doesn't expose the owner.
    $namedEnvs = @()
    if ($nameTokens.Count -gt 0) {
        $namedEnvs = @($matchEnvs | Where-Object {
            $d = [string]$_.DisplayName
            if (-not $d) { return $false }
            foreach ($t in $nameTokens) {
                if ($d.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
            }
            return $false
        })
    }
    Trace ""name-matched envs: $($namedEnvs.Count)""

    # Merge exact-name+owned+token-name, dedupe. Because $matchEnvs is Developer
    # first, exact friendly-name matches in the user's Developer environment win.
    $candidates = @()
    $seenIds    = @{}
    foreach ($e in $exactDisplayNameEnvs) {
        $id = [string]$e.EnvironmentName
        if ($id -and -not $seenIds.ContainsKey($id)) { $seenIds[$id] = $true; $candidates += $e }
    }
    foreach ($e in $ownedEnvs) {
        $id = [string]$e.EnvironmentName
        if ($id -and -not $seenIds.ContainsKey($id)) { $seenIds[$id] = $true; $candidates += $e }
    }
    foreach ($e in $namedEnvs) {
        $id = [string]$e.EnvironmentName
        if ($id -and -not $seenIds.ContainsKey($id)) { $seenIds[$id] = $true; $candidates += $e }
    }
    Trace ""candidate envs (developer-first exact+owned+named, deduped): $($candidates.Count)""
    foreach ($e in $candidates) {
        $iu = $null
        try { $iu = $e.Internal.properties.linkedEnvironmentMetadata.instanceUrl } catch {}
        Trace ""  candidate: displayName='$($e.DisplayName)' id='$($e.EnvironmentName)' hasInstanceUrl=$([bool]$iu)""
    }

    # Pick the first candidate with a Dataverse instanceUrl (real DB).
    $target = $null
    foreach ($e in $candidates) {
        $iu = $null
        try { $iu = $e.Internal.properties.linkedEnvironmentMetadata.instanceUrl } catch {}
        if ($iu) { $target = $e; break }
    }
    }

    if (-not $target) {
        $hint = ''
        if ($candidates.Count -gt 0) {
            $names = @($candidates | ForEach-Object { ""'$($_.DisplayName)'"" }) -join ', '
            $hint  = "" Found $($candidates.Count) candidate env(s) but none have a Dataverse database: $names.""
        }
        Emit 'NoOwnedEnvironment' (
            ""Couldn't find a Power Platform environment with a Dataverse database that you own."" + $hint +
            "" Create a free Developer Environment at https://make.powerapps.com/environments and re-try, "" +
            ""or use the Manual setup guide."" ) $null $null $null $null
    }

    $envId      = [string]$target.EnvironmentName
    $envDisplay = [string]$target.DisplayName
    $instance   = ([string]$target.Internal.properties.linkedEnvironmentMetadata.instanceUrl).TrimEnd('/')
    Trace ""target env: id=$envId display='$envDisplay' instance=$instance""
    Report ""Selected '$envDisplay'. Preparing automatic import...""
    }

    # Let Power Platform CLI own Dataverse auth/import instead of borrowing a
    # Dataverse token from the PowerApps PowerShell module. The latter is blocked
    # in some locked-down tenants with AADSTS65002.
    $pacPath = $null
    $pacCmd = Get-Command pac -ErrorAction SilentlyContinue
    if ($pacCmd -and $pacCmd.Source) { $pacPath = [string]$pacCmd.Source }
    if (-not $pacPath -and $env:USERPROFILE) {
        $dotnetToolPac = Join-Path $env:USERPROFILE '.dotnet\tools\pac.exe'
        if (Test-Path -LiteralPath $dotnetToolPac) { $pacPath = $dotnetToolPac }
    }
    if (-not $pacPath -and $env:LOCALAPPDATA) {
        $msiPacCmd = Join-Path $env:LOCALAPPDATA 'Microsoft\PowerAppsCLI\pac.cmd'
        if (Test-Path -LiteralPath $msiPacCmd) { $pacPath = $msiPacCmd }
    }
    if (-not $pacPath -and $env:LOCALAPPDATA) {
        $msiPacLauncher = Join-Path $env:LOCALAPPDATA 'Microsoft\PowerAppsCLI\pac.launcher.exe'
        if (Test-Path -LiteralPath $msiPacLauncher) { $pacPath = $msiPacLauncher }
    }
    if (-not $pacPath -and $env:LOCALAPPDATA) {
        $msiPacExe = Join-Path $env:LOCALAPPDATA 'Microsoft\PowerAppsCLI\pac.exe'
        if (Test-Path -LiteralPath $msiPacExe) { $pacPath = $msiPacExe }
    }
    if (-not $pacPath) {
        Emit 'OtherError' (
            ""Automatic Power Automate import is not available on this PC. Install the Microsoft Power Platform CLI to enable it."" + [Environment]::NewLine +
            ""Opening Power Automate solutions page for '$envDisplay' so you can import OofManager-CloudSchedule.zip manually.""
        ) $envId $envDisplay $instance $null
    }
    Trace ""pac path: $pacPath""

    if (-not (Test-Path -LiteralPath $zipPath)) {
        Emit 'OtherError' ""Solution zip missing: $zipPath"" $envId $envDisplay $instance $null
    }
    $zipBytes = [System.IO.File]::ReadAllBytes($zipPath).Length
    Trace ""zip size: $zipBytes bytes""

    function Invoke-Pac($stepName, [string[]]$arguments) {
        Trace (""pac "" + ($arguments -join ' '))
        & $pacPath @arguments 2>&1 | ForEach-Object {
            $line = [string]$_
            Trace (""pac> "" + $line)
            if ($line -match 'Connected to\.\.\.') { Report $line }
            elseif ($line -match 'Solution Importing') { Report ""Importing solution '$solName' into '$envDisplay'..."" }
            elseif ($line -match 'Solution Imported successfully') { Report 'Solution imported. Publishing customizations...' }
            elseif ($line -match 'Publishing All Customizations') { Report 'Publishing Power Automate customizations...' }
            elseif ($line -match 'Published All Customizations') { Report 'Published customizations. Finishing up...' }
            Write-Host $line
        }
        $code = $LASTEXITCODE
        Trace ""pac $stepName exit=$code""
        return $code
    }

    $authWhoCode = Invoke-Pac 'auth who' @('auth', 'who')
    if ($authWhoCode -ne 0) {
        # Microsoft corporate accounts use pac's WAM/OperatingSystem profile;
        # creating a browser auth profile can be rejected even when WAM works.
        Trace 'no active pac auth profile; attempting pac auth create'
        $authCode = Invoke-Pac 'auth create' @('auth', 'create', '--environment', $instance)
        if ($authCode -ne 0) {
            Emit 'SignInFailed' (""Automatic Power Automate sign-in failed. Open Power Automate once in your browser, then try again."" ) $envId $envDisplay $instance $null
        }
    } else {
        Trace 'using existing pac auth profile'
        Report 'Signed in. Uploading solution to Power Automate...'
    }

    $importArgs = @(
        'solution', 'import',
        '--path', $zipPath,
        '--environment', $instance,
        '--publish-changes',
        '--max-async-wait-time', '10'
    )
    if ($forceFlag -eq '1') {
        $importArgs += '--force-overwrite'
    }

    $importCode = Invoke-Pac 'solution import' $importArgs
    if ($importCode -ne 0) {
        Emit 'ImportFailed' (""Power Automate rejected the solution import for '$solName' in '$envDisplay'."" ) $envId $envDisplay $instance $null
    }

    Report 'Import complete.'
    Emit 'Success' (""Imported solution '"" + $solName + ""' into '"" + $envDisplay + ""'."" ) $envId $envDisplay $instance $workflowId

} catch {
    Trace ""OUTER CATCH: $($_.Exception.Message)""
    Emit 'OtherError' $_.Exception.Message $null $null $null $null
}
";
}
