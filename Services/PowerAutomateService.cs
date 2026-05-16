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

    public Task<PowerAutomateStatusResult> GetOofManagerFlowStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default)
        => RunStatusAsync(upnHint, displayNameHint, expectedFlowDisplayName, ct);

    public Task<PowerAutomateResult> DisableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default)
        => RunOperationAsync("Disable-Flow", upnHint, displayNameHint, expectedFlowDisplayName, ct);

    public Task<PowerAutomateResult> EnableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default)
        => RunOperationAsync("Enable-Flow", upnHint, displayNameHint, expectedFlowDisplayName, ct);

    public async Task<CloudScheduleImportResult> ImportCloudScheduleSolutionAsync(
        string solutionZipPath,
        string solutionUniqueName,
        Guid workflowId,
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        CancellationToken ct = default)
    {
        if (!File.Exists(solutionZipPath))
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.OtherError,
                $"Solution zip not found at '{solutionZipPath}'.",
                null, null, null, null);
        }

        var envVars = new Dictionary<string, string>
        {
            ["OOFMGR_PA_UPN"]          = upnHint ?? string.Empty,
            ["OOFMGR_PA_DISPLAYNAME"]  = displayNameHint ?? string.Empty,
            ["OOFMGR_PA_ZIPPATH"]      = solutionZipPath,
            ["OOFMGR_PA_SOLNAME"]      = solutionUniqueName,
            ["OOFMGR_PA_WORKFLOWID"]   = workflowId.ToString("D"),
            ["OOFMGR_PA_FORCE"]        = forceOverwrite ? "1" : "0",
        };

        // Import + poll can take longer than a flow toggle: ImportSolutionAsync
        // averages 30-90 s on a fresh dev env, plus 50 s for the env listing on
        // big tenants and 10 s for first-time sign-in. 7 min outer cap leaves
        // generous headroom over the 5-min PS-side poll deadline.
        var json = await RunPowerShellChildAsync(
            BuildImportScript(),
            envVars,
            timeout: TimeSpan.FromMinutes(7),
            ct).ConfigureAwait(false);

        return ParseImportResult(json.Json, json.ExitCode);
    }

    private async Task<PowerAutomateStatusResult> RunStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct)
    {
        var json = await RunFlowScriptAsync("Get-Status", upnHint, displayNameHint, expectedFlowDisplayName, ct).ConfigureAwait(false);
        return ParseStatusResult(json.Json, json.ExitCode);
    }

    private async Task<PowerAutomateResult> RunOperationAsync(string verb, string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct)
    {
        var json = await RunFlowScriptAsync(verb, upnHint, displayNameHint, expectedFlowDisplayName, ct).ConfigureAwait(false);
        return ParseResult(json.Json, json.ExitCode);
    }

    private Task<ChildResult> RunFlowScriptAsync(string verb, string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct)
    {
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
        };
        return RunPowerShellChildAsync(
            BuildScript(),
            envVars,
            timeout: TimeSpan.FromMinutes(5),
            ct);
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
        CancellationToken ct)
    {
        var modulesDir = Path.Combine(AppContext.BaseDirectory, "Modules");
        var psFile = Path.Combine(Path.GetTempPath(), $"oofmgr-pa-{Guid.NewGuid():N}.ps1");
        var resultFile = Path.Combine(Path.GetTempPath(), $"oofmgr-pa-result-{Guid.NewGuid():N}.json");
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
            foreach (var kv in envVars)
            {
                psi.EnvironmentVariables[kv.Key] = kv.Value;
            }

            using var proc = new Process { StartInfo = psi };
            proc.Start();

            using (ct.Register(() => { try { if (!proc.HasExited) proc.Kill(); } catch { } }))
            {
                var exited = await Task.Run(() => proc.WaitForExit((int)timeout.TotalMilliseconds), ct).ConfigureAwait(false);
                if (!exited)
                {
                    try { proc.Kill(); } catch { }
                    return new ChildResult(null, -1);
                }
            }

            string? resultJson = null;
            try { if (File.Exists(resultFile)) resultJson = File.ReadAllText(resultFile, Encoding.UTF8); } catch { }
            return new ChildResult(resultJson, proc.ExitCode);
        }
        finally
        {
            try { File.Delete(psFile); } catch { /* best-effort */ }
            try { File.Delete(resultFile); } catch { /* best-effort */ }
        }
    }

    private static CloudScheduleImportResult ParseImportResult(string? resultJson, int exitCode)
    {
        if (exitCode == -1 && resultJson is null)
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.SignInFailed,
                "Power Automate sign-in / import did not complete within 7 minutes. Falling back to the manual setup flow.",
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

            var outcome = outcomeStr switch
            {
                "Success" => PowerAutomateOutcome.Success,
                "NoFlowFound" => PowerAutomateOutcome.NoFlowFound,
                "SignInFailed" => PowerAutomateOutcome.SignInFailed,
                "SolutionAwareBlocked" => PowerAutomateOutcome.SolutionAwareBlocked,
                _ => PowerAutomateOutcome.OtherError,
            };
            return new PowerAutomateResult(outcome, message, flows);
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
            return new PowerAutomateStatusResult(outcome, state, message, flows);
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
$resultFile      = $env:OOFMGR_PA_RESULTFILE

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
Trace ""start: pid=$PID verb=$verb prefix='$prefix' flowDisplayName='$flowDisplayName' upn='$upn' displayName='$displayName' modulesDir='$modulesDir'""

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

function Emit($outcome, $message, $flows, $state = $null) {
    $obj = [ordered]@{
        Outcome = $outcome
        Message = $message
        Flows   = @($flows)
        State   = $state
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
    Trace 'before Import-Module Microsoft.PowerApps.PowerShell'
    Import-Module Microsoft.PowerApps.PowerShell -DisableNameChecking -ErrorAction Stop | Out-Null
    Trace 'after  Import-Module Microsoft.PowerApps.PowerShell'

    # Sign in. With a valid cached refresh token (under
    # %LOCALAPPDATA%\Microsoft\PowerAppsCli) this returns silently; otherwise
    # ADAL pops its sign-in dialog as a separate top-level window.
    try {
        if ($upn) {
            Trace ""before Add-PowerAppsAccount -Username $upn""
            Add-PowerAppsAccount -Endpoint prod -Username $upn -ErrorAction Stop | Out-Null
        } else {
            Trace 'before Add-PowerAppsAccount (no upn)'
            Add-PowerAppsAccount -Endpoint prod -ErrorAction Stop | Out-Null
        }
        Trace 'after  Add-PowerAppsAccount'
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
        foreach ($f in $flows) {
            $dn = [string]$f.DisplayName
            if ($dn -and $flowDisplayName -and ($dn -ieq $flowDisplayName)) {
                $script:matched += [PSCustomObject]@{
                    EnvironmentName = $envName
                    FlowName        = $f.FlowName
                    DisplayName     = $dn
                    Enabled         = $f.Enabled
                }
                Trace ""matched flow by exact displayName: $dn enabled=$($f.Enabled)""
            }
        }
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

    function Get-DeveloperAdminEnvsFast() {
        $developerEnvs = @()
        try {
            # This is the first-pass narrow query: unlike Get-PowerAppEnvironment,
            # it asks the service for Developer envs before we ever enumerate the
            # whole tenant. Some tenants/API versions may reject this $filter; in
            # that case we trace and let the existing full-scan fallback handle it.
            $developerEnvUri = ""https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/environments?`$expand=permissions&`$filter=properties/environmentSku%20eq%20%27Developer%27&api-version={apiVersion}""
            Trace 'before fast developer-env query'
            $developerResult = InvokeApi -Method GET -Route $developerEnvUri -ErrorAction Stop
            $developerEnvs = @($developerResult.value | ForEach-Object { Convert-RawEnvironment $_ } | Where-Object { $_ -and (Test-DeveloperEnvironment $_) })
            Trace ""after  fast developer-env query: count=$($developerEnvs.Count)""
        } catch {
            Trace ""fast developer-env query failed: $($_.Exception.Message)""
            return @()
        }

        $adminDeveloperEnvs = @($developerEnvs | Where-Object { Test-AdminEnvironmentForCurrentUser $_ })
        Trace ""admin+developer envs from fast query: count=$($adminDeveloperEnvs.Count)""

        if ($adminDeveloperEnvs.Count -gt 0) { return $adminDeveloperEnvs }

        # Permission shape is not consistent across tenants. If the service did
        # return Developer envs but no obvious admin marker, still scan those
        # Developer envs first; this keeps the first pass narrow without risking
        # a false negative caused by a missing/renamed permission key.
        if ($developerEnvs.Count -gt 0) {
            Trace 'no explicit admin marker found; scanning developer envs anyway before full fallback'
            return $developerEnvs
        }

        return @()
    }

    $fastDeveloperAdminEnvs = @(Get-DeveloperAdminEnvsFast)
    $fastDeveloperScanned = 0
    foreach ($e in $fastDeveloperAdminEnvs) {
        if ($matched.Count -gt 0) { break }
        if ((Get-Date) -gt $deadline) { break }
        Trace ""fast developer/admin env[$fastDeveloperScanned] displayName='$($e.DisplayName)' id='$($e.EnvironmentName)'""
        Scan-Env $e.EnvironmentName ""fast-developer-admin-env[$fastDeveloperScanned]""
        $fastDeveloperScanned++
    }
    if ($matched.Count -gt 0) {
        Trace 'matched in fast developer/admin pass; skipping full environment scan'
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
    Trace 'before Get-PowerAppEnvironment (all)'
    $allEnvs = @(Get-PowerAppEnvironment -ErrorAction SilentlyContinue)
    Trace ""after  Get-PowerAppEnvironment: count=$($allEnvs.Count)""

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

    # Merge owned + named, dedupe, keep order (owned-first).
    $priorityEnvs = New-Object 'System.Collections.Generic.List[object]'
    $seenIds      = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $ownedEnvs) {
        if ($e.EnvironmentName -and $seenIds.Add([string]$e.EnvironmentName)) { $priorityEnvs.Add($e) | Out-Null }
    }
    foreach ($e in $namedEnvs) {
        if ($e.EnvironmentName -and $seenIds.Add([string]$e.EnvironmentName)) { $priorityEnvs.Add($e) | Out-Null }
    }
    Trace ""priority envs (owned+named, deduped): count=$($priorityEnvs.Count)""
    $i = 0
    foreach ($e in $priorityEnvs) {
        if ($matched.Count -gt 0) { break }
        if ((Get-Date) -gt $deadline) { break }
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
            $freshState = Read-CurrentEnabled $f
            if ($freshState.Known) {
                $state = if ($freshState.Enabled) { 'On' } else { 'Off' }
                $lowerState = $state.ToLowerInvariant()
                Emit 'Success' ('Power Automate flow is ' + $lowerState) @($f.DisplayName) $state
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
            $enabledProp = $f.PSObject.Properties['Enabled']
            if ($enabledProp -ne $null) {
                $enabledType = if ($enabledProp.Value -ne $null) { $enabledProp.Value.GetType().FullName } else { '<null>' }
                Trace ""reported flow state before ${verb}: $($f.DisplayName) enabledRaw='$($enabledProp.Value)' type='$enabledType' desiredEnabled=$desiredEnabled""
            }

            $freshState = Read-CurrentEnabled $f
            if ($freshState.Known -and $freshState.Enabled -eq $desiredEnabled) {
                Trace ""skip $verb; fresh state already desired: $($f.DisplayName)""
                $alreadyDesired += $f.DisplayName
                continue
            }

            Trace ""before $verb env=$($f.EnvironmentName) flow=$($f.FlowName)""
            if ($f.EnvironmentName) {
                & $verb -EnvironmentName $f.EnvironmentName -FlowName $f.FlowName -ErrorAction Stop | Out-Null
            } else {
                & $verb -FlowName $f.FlowName -ErrorAction Stop | Out-Null
            }
            Trace ""after  $verb OK: $($f.DisplayName)""
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
    Emit 'Success' $msg @($changed + $alreadyDesired)

} catch {
    Trace ""OUTER CATCH: $($_.Exception.Message)""
    Emit 'OtherError' $_.Exception.Message @()
}
";

    /// <summary>
    /// Payload for <see cref="ImportCloudScheduleSolutionAsync"/>. Auths via the
    /// bundled Microsoft.PowerApps.PowerShell module the same way the
    /// toggle script does, picks the env whose principal-owner email
    /// matches the signed-in UPN (Dataverse solution import requires a real
    /// Dataverse-backed env and the user must own / co-own it), then talks
    /// directly to the Dataverse Web API at
    /// <c>{instanceUrl}/api/data/v9.2/</c> for:
    ///   1. GET /solutions?$filter=uniquename eq '…' — existence check.
    ///   2. POST /ImportSolutionAsync — base64-encoded zip + ImportJobId.
    ///   3. Poll /asyncoperations({id}) until statecode == 3 (Completed).
    ///   4. POST /PublishAllXml — required so the imported flow's trigger
    ///      bindings get rebuilt; without this the flow shows up as
    ///      "Cannot publish" in the Maker UI.
    ///   5. Echo back env id + instanceUrl + workflow id so the C# layer
    ///      can build a https://make.powerautomate.com/environments/{env}/
    ///      flows/{workflowId}/details deep link.
    /// The Dataverse token is obtained via <c>Get-JwtToken -Audience
    /// "$instanceUrl/"</c> — the AuthModule publicly exports that helper
    /// and it silently mints an org-scoped token from the cached MSAL
    /// refresh token populated by <c>Add-PowerAppsAccount</c>.
    /// </summary>
    private static string BuildImportScript() => @"
$ErrorActionPreference = 'Stop'
$modulesDir   = $env:OOFMGR_PA_MODULES
$upn          = $env:OOFMGR_PA_UPN
$displayName  = $env:OOFMGR_PA_DISPLAYNAME
$resultFile   = $env:OOFMGR_PA_RESULTFILE
$zipPath      = $env:OOFMGR_PA_ZIPPATH
$solName      = $env:OOFMGR_PA_SOLNAME
$workflowId   = $env:OOFMGR_PA_WORKFLOWID
$forceFlag    = $env:OOFMGR_PA_FORCE

$debugLog = Join-Path $env:TEMP 'oofmgr-pa-debug.log'
try { Remove-Item -LiteralPath $debugLog -ErrorAction SilentlyContinue } catch {}
function Trace($m) {
    try {
        $line = ('{0}  {1}' -f (Get-Date -Format o), $m)
        Add-Content -LiteralPath $debugLog -Value $line -Encoding UTF8
    } catch {}
}
Trace ""start: pid=$PID op=import upn='$upn' solName='$solName' workflowId='$workflowId' force='$forceFlag' zip='$zipPath'""

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
    if ($modulesDir -and (Test-Path -LiteralPath $modulesDir)) {
        $env:PSModulePath = $modulesDir + ';' + $env:PSModulePath
    }
    $ProgressPreference = 'SilentlyContinue'
    $WarningPreference  = 'SilentlyContinue'

    Import-Module Microsoft.PowerApps.PowerShell -DisableNameChecking -ErrorAction Stop | Out-Null
    Trace 'module loaded'

    try {
        if ($upn) { Add-PowerAppsAccount -Endpoint prod -Username $upn -ErrorAction Stop | Out-Null }
        else      { Add-PowerAppsAccount -Endpoint prod -ErrorAction Stop | Out-Null }
        Trace 'signed in'
    } catch {
        Trace ""sign-in FAILED: $($_.Exception.Message)""
        Emit 'SignInFailed' (""Sign-in to Power Automate failed: "" + $_.Exception.Message) $null $null $null $null
    }

    Trace 'before Get-PowerAppEnvironment'
    $allEnvs = @(Get-PowerAppEnvironment -ErrorAction SilentlyContinue)
    Trace ""env count: $($allEnvs.Count)""

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

    # 1) Strict UPN match — usually 0-3 envs on a Microsoft tenant.
    $ownedEnvs = @()
    if ($upn) {
        $ownedEnvs = @($allEnvs | Where-Object {
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

    # 2) Display-name match — picks up 'Sandy Sun's Environment' style
    #    dev envs whose principal field doesn't expose the owner.
    $namedEnvs = @()
    if ($nameTokens.Count -gt 0) {
        $namedEnvs = @($allEnvs | Where-Object {
            $d = [string]$_.DisplayName
            if (-not $d) { return $false }
            foreach ($t in $nameTokens) {
                if ($d.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
            }
            return $false
        })
    }
    Trace ""name-matched envs: $($namedEnvs.Count)""

    # Merge owned+named, dedupe (owned first).
    $candidates = New-Object 'System.Collections.Generic.List[object]'
    $seenIds    = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $ownedEnvs) {
        if ($e.EnvironmentName -and $seenIds.Add([string]$e.EnvironmentName)) { $candidates.Add($e) | Out-Null }
    }
    foreach ($e in $namedEnvs) {
        if ($e.EnvironmentName -and $seenIds.Add([string]$e.EnvironmentName)) { $candidates.Add($e) | Out-Null }
    }
    Trace ""candidate envs (owned+named, deduped): $($candidates.Count)""
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

    # Get the Dataverse-scoped bearer token. The PowerApps module's
    # AuthModule publicly exports Get-JwtToken; calling it with -Audience
    # set to the instance URL triggers an internal silent token refresh
    # against MSAL (the user already signed in above for the Power Apps
    # audience), so no second prompt.
    Trace 'before Get-JwtToken'
    $dvToken = $null
    try {
        $dvToken = Get-JwtToken -Audience ($instance + '/')
    } catch {
        $exMsg = $_.Exception.Message
        Trace ""Get-JwtToken FAILED: $exMsg""
        # AADSTS65002 = first-party app preauthorization missing. Locked-down
        # tenants (e.g. Microsoft corp) block the bundled PowerApps PowerShell
        # client ID from minting Dataverse tokens. Surface a distinct outcome
        # so the UI can route to a friendly fallback instead of a wall of
        # AAD error text.
        if ($exMsg -match 'AADSTS65002') {
            Emit 'TenantBlocked' (
                ""Your Microsoft tenant doesn't allow auto-import via the PowerApps PowerShell module (AADSTS65002)."" + [Environment]::NewLine +
                ""Opening Power Automate solutions page for '$envDisplay' — confirm this is the environment named after you, then click 'Import solution' and pick OofManager-CloudSchedule.zip from your Desktop.""
            ) $envId $envDisplay $instance $null
        }
        Emit 'OtherError' (""Could not get Dataverse token for $instance : "" + $exMsg) $envId $envDisplay $instance $null
    }
    if (-not $dvToken) {
        Emit 'OtherError' ""Empty Dataverse token for $instance"" $envId $envDisplay $instance $null
    }
    Trace ""got Dataverse token (len=$($dvToken.Length))""

    $apiBase = $instance + '/api/data/v9.2'
    $headers = @{
        'Authorization'    = ""Bearer $dvToken""
        'Accept'           = 'application/json'
        'OData-MaxVersion' = '4.0'
        'OData-Version'    = '4.0'
    }

    # Existence check — only when not forced. If the same uniquename
    # already exists, bail with AlreadyExists so the C# layer can confirm
    # with the user before clobbering it.
    if ($forceFlag -ne '1') {
        try {
            $checkUrl = $apiBase + ""/solutions?`$filter=uniquename eq '"" + $solName + ""'&`$select=solutionid,friendlyname,version""
            Trace ""GET $checkUrl""
            $existing = Invoke-RestMethod -Method Get -Uri $checkUrl -Headers $headers
            if ($existing.value -and $existing.value.Count -gt 0) {
                $v = $existing.value[0]
                Trace ""solution exists: id=$($v.solutionid) version=$($v.version)""
                Emit 'AlreadyExists' (
                    ""Solution '"" + $solName + ""' already exists in '"" + $envDisplay + ""' (v"" + $v.version + "")."" )                   $envId $envDisplay $instance $null
            }
        } catch {
            Trace ""existence check FAILED (non-fatal): $($_.Exception.Message)""
        }
    }

    # Read + base64 the zip.
    if (-not (Test-Path -LiteralPath $zipPath)) {
        Emit 'OtherError' ""Solution zip missing: $zipPath"" $envId $envDisplay $instance $null
    }
    $bytes = [System.IO.File]::ReadAllBytes($zipPath)
    $b64   = [Convert]::ToBase64String($bytes)
    $importJobId = [Guid]::NewGuid().ToString()
    Trace ""zip size: $($bytes.Length) bytes, importJobId=$importJobId""

    # Kick off the async import.
    $importBody = @{
        OverwriteUnmanagedCustomizations = $true
        PublishWorkflows                 = $true
        CustomizationFile                = $b64
        ImportJobId                      = $importJobId
    } | ConvertTo-Json -Compress

    try {
        Trace 'POST /ImportSolutionAsync'
        $importResp = Invoke-RestMethod -Method Post -Uri ($apiBase + '/ImportSolutionAsync') -Headers $headers -Body $importBody -ContentType 'application/json'
    } catch {
        Trace ""ImportSolutionAsync FAILED: $($_.Exception.Message)""
        Emit 'ImportFailed' (""ImportSolutionAsync failed: "" + $_.Exception.Message) $envId $envDisplay $instance $null
    }

    $asyncOpId = $null
    if ($importResp -and $importResp.AsyncOperationId) {
        $asyncOpId = [string]$importResp.AsyncOperationId
    } elseif ($importResp -and $importResp.asyncoperationid) {
        $asyncOpId = [string]$importResp.asyncoperationid
    }
    if (-not $asyncOpId) {
        Trace ""no asyncoperationid in response: $($importResp | ConvertTo-Json -Compress -Depth 4)""
        Emit 'ImportFailed' ""ImportSolutionAsync returned no AsyncOperationId."" $envId $envDisplay $instance $null
    }
    Trace ""asyncOpId=$asyncOpId""

    # Poll for completion. statecode 3 == Completed; statuscode 30 ==
    # Succeeded, anything else is a failure to surface to the user.
    $pollDeadline = (Get-Date).AddMinutes(5)
    $async = $null
    while ((Get-Date) -lt $pollDeadline) {
        Start-Sleep -Seconds 4
        try {
            $async = Invoke-RestMethod -Method Get -Uri ($apiBase + ""/asyncoperations($asyncOpId)?`$select=statecode,statuscode,message,friendlymessage"") -Headers $headers
        } catch {
            Trace ""poll FAILED (will retry): $($_.Exception.Message)""
            continue
        }
        Trace ""poll: statecode=$($async.statecode) statuscode=$($async.statuscode)""
        if ($async.statecode -eq 3) { break }
    }
    if (-not $async -or $async.statecode -ne 3) {
        Emit 'ImportFailed' ""Solution import did not complete within 5 minutes (last statecode=$($async.statecode))."" $envId $envDisplay $instance $null
    }
    if ($async.statuscode -ne 30) {
        $errMsg = $async.friendlymessage; if (-not $errMsg) { $errMsg = $async.message }
        Emit 'ImportFailed' (""Solution import failed (statuscode=$($async.statuscode)): "" + $errMsg) $envId $envDisplay $instance $null
    }
    Trace 'import succeeded, publishing'

    # Publish all customizations so the imported workflow's trigger
    # bindings become active. Without this the flow shows in the Maker
    # UI but the trigger schedule won't fire.
    try {
        Invoke-RestMethod -Method Post -Uri ($apiBase + '/PublishAllXml') -Headers $headers -Body '{}' -ContentType 'application/json' | Out-Null
        Trace 'PublishAllXml ok'
    } catch {
        Trace ""PublishAllXml FAILED (non-fatal): $($_.Exception.Message)""
    }

    Emit 'Success' (""Imported solution '"" + $solName + ""' into '"" + $envDisplay + ""'."" ) $envId $envDisplay $instance $workflowId

} catch {
    Trace ""OUTER CATCH: $($_.Exception.Message)""
    Emit 'OtherError' $_.Exception.Message $null $null $null $null
}
";
}
