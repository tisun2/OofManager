using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using OofManager.Wpf.Models;

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
    private const string FlowDisplayNamePrefix = "OofManager Weekly Schedule";
    private const string ImportCacheUpnKey = "PowerAutomate.Import.Environment.Upn";
    private const string ImportCacheEnvironmentIdKey = "PowerAutomate.Import.Environment.Id";
    private const string ImportCacheEnvironmentDisplayNameKey = "PowerAutomate.Import.Environment.DisplayName";
    private const string ImportCacheInstanceUrlKey = "PowerAutomate.Import.Environment.InstanceUrl";
    private const string FlowCacheUpnKey = "PowerAutomate.Flow.Upn";
    private const string FlowCacheEnvironmentIdKey = "PowerAutomate.Flow.Environment.Id";
    private const string FlowCacheDisplayNameKey = "PowerAutomate.Flow.DisplayName";
    private const string FlowCacheNameKey = "PowerAutomate.Flow.Name";
    private const string FlowReferencesCacheV2Key = "PowerAutomate.Flow.References.V2";
    private const string ConnectionCacheV1Key = "PowerAutomate.Connection.Cache.V1";

    private readonly IPreferencesService _prefs;
    private readonly object _flowCacheGate = new();

    public PowerAutomateService(IPreferencesService prefs)
    {
        _prefs = prefs;
    }

    public PowerAutomateFlowReference? TryGetCachedFlowReference(string? upnHint, string expectedFlowDisplayName)
    {
        var cachedEnvironment = GetCachedImportEnvironment(upnHint);
        var cached = GetCachedFlowReference(upnHint, expectedFlowDisplayName, cachedEnvironment?.EnvironmentId);
        return cached is null
            ? null
            : new PowerAutomateFlowReference(cached.EnvironmentName, cached.FlowName, cached.DisplayName);
    }

    public Task<PowerAutomateStatusResult> GetOofManagerFlowStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default, IProgress<string>? progress = null)
        => RunStatusAsync(upnHint, displayNameHint, expectedFlowDisplayName, ct, progress);

    public Task<PowerAutomateFlowBatchResult> GetOofManagerFlowsStatusBatchAsync(string? upnHint, string? displayNameHint, IReadOnlyList<string> expectedFlowDisplayNames, CancellationToken ct = default, IProgress<string>? progress = null)
        => RunFlowsBatchVerbAsync("Get-Flows-Status", upnHint, displayNameHint, expectedFlowDisplayNames, progress, ct, interactive: false);

    public Task<PowerAutomateResult> DisableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default)
        => RunOperationAsync("Disable-Flow", upnHint, displayNameHint, expectedFlowDisplayName, progress, ct);

    public Task<PowerAutomateResult> EnableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default)
        => RunOperationAsync("Enable-Flow", upnHint, displayNameHint, expectedFlowDisplayName, progress, ct);

    public Task<PowerAutomateFlowBatchResult> SetOofManagerFlowsEnabledBatchAsync(string? upnHint, string? displayNameHint, IReadOnlyList<string> expectedFlowDisplayNames, bool enabled, IProgress<string>? progress = null, CancellationToken ct = default)
        => RunFlowsBatchVerbAsync(enabled ? "Enable-Flows-Batch" : "Disable-Flows-Batch", upnHint, displayNameHint, expectedFlowDisplayNames, progress, ct, interactive: true);

    public Task<PowerAutomateFlowBatchResult> EnsureOofManagerFlowsOnAsync(string? upnHint, string? displayNameHint, IReadOnlyList<string> expectedFlowDisplayNames, IProgress<string>? progress = null, CancellationToken ct = default)
        => RunFlowsBatchVerbAsync("Ensure-Flows-On", upnHint, displayNameHint, expectedFlowDisplayNames, progress, ct, interactive: true);

    public async Task<CloudScheduleDefinitionResult> GetCloudScheduleDefinitionAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default)
    {
        var child = await RunFlowScriptAsync("Get-Definition", upnHint, displayNameHint, expectedFlowDisplayName, ct, progress).ConfigureAwait(false);
        return ParseDefinitionResult(child.Json, child.ExitCode, expectedFlowDisplayName);
    }

    private static CloudScheduleDefinitionResult ParseDefinitionResult(string? resultJson, int exitCode, string expectedFlowDisplayName)
    {
        if (exitCode == -1 && resultJson is null)
        {
            return new CloudScheduleDefinitionResult(PowerAutomateOutcome.SignInFailed,
                "Power Automate sign-in did not complete within 5 minutes.", expectedFlowDisplayName, null, null, null, null);
        }
        if (string.IsNullOrWhiteSpace(resultJson))
        {
            return new CloudScheduleDefinitionResult(PowerAutomateOutcome.OtherError,
                $"PowerShell child exited with code {exitCode} but produced no result file.", expectedFlowDisplayName, null, null, null, null);
        }
        try
        {
            using var doc = JsonDocument.Parse(resultJson!);
            var root = doc.RootElement;
            var outcomeStr = root.TryGetProperty("Outcome", out var oEl) ? oEl.GetString() : null;
            var message = root.TryGetProperty("Message", out var mEl) ? mEl.GetString() ?? "" : "";
            var outcome = outcomeStr switch
            {
                "Success" => PowerAutomateOutcome.Success,
                "NoFlowFound" => PowerAutomateOutcome.NoFlowFound,
                "SignInFailed" => PowerAutomateOutcome.SignInFailed,
                "SolutionAwareBlocked" => PowerAutomateOutcome.SolutionAwareBlocked,
                _ => PowerAutomateOutcome.OtherError,
            };

            string? flowDisplayName = null;
            if (root.TryGetProperty("Flows", out var fEl) && fEl.ValueKind == JsonValueKind.Array && fEl.GetArrayLength() > 0)
                flowDisplayName = fEl[0].GetString();

            if (outcome != PowerAutomateOutcome.Success)
                return new CloudScheduleDefinitionResult(outcome, message, flowDisplayName, null, null, null, null);

            if (!root.TryGetProperty("Definition", out var defEl) || defEl.ValueKind != JsonValueKind.Object)
                return new CloudScheduleDefinitionResult(PowerAutomateOutcome.OtherError,
                    "Cloud flow definition was empty.", flowDisplayName, null, null, null, null);

            // Trigger days are independent of the workdays in the sidecar:
            // daily flows also restore OOF after non-workday overrides.
            var triggerDays = new List<DayOfWeek>();
            int? triggerHour = null, triggerMinute = null;
            string? triggerTz = null;
            DateTimeOffset? triggerStartTimeUtc = null;
            if (defEl.TryGetProperty("triggers", out var trigsEl) && trigsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var trigProp in trigsEl.EnumerateObject())
                {
                    if (!trigProp.Value.TryGetProperty("recurrence", out var recEl) || recEl.ValueKind != JsonValueKind.Object)
                        continue;
                    if (recEl.TryGetProperty("timeZone", out var tzEl) && tzEl.ValueKind == JsonValueKind.String)
                        triggerTz = tzEl.GetString();

                    if (recEl.TryGetProperty("frequency", out var freqEl) && freqEl.ValueKind == JsonValueKind.String &&
                        string.Equals(freqEl.GetString(), "Day", StringComparison.OrdinalIgnoreCase) &&
                        recEl.TryGetProperty("interval", out var intervalEl) && intervalEl.TryGetInt32(out var interval) && interval == 1)
                    {
                        triggerDays.AddRange((DayOfWeek[])Enum.GetValues(typeof(DayOfWeek)));
                    }

                    // Vacation flows: one-shot Recurrence anchored at a
                    // specific UTC instant via 'startTime'. Parse it so the
                    // Manual-side "Compare with cloud" can verify the
                    // deployed flow fires at the user's planned vacation
                    // start / end.
                    if (recEl.TryGetProperty("startTime", out var stEl) && stEl.ValueKind == JsonValueKind.String &&
                        DateTimeOffset.TryParse(stEl.GetString(), System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out var parsedSt))
                    {
                        triggerStartTimeUtc = parsedSt;
                    }

                    if (recEl.TryGetProperty("schedule", out var schedEl) && schedEl.ValueKind == JsonValueKind.Object)
                    {
                        if (triggerDays.Count == 0 && schedEl.TryGetProperty("weekDays", out var wdEl) && wdEl.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var d in wdEl.EnumerateArray())
                            {
                                if (d.ValueKind == JsonValueKind.String && Enum.TryParse<DayOfWeek>(d.GetString(), true, out var dow))
                                    triggerDays.Add(dow);
                            }
                        }
                        if (schedEl.TryGetProperty("hours", out var hEl) && hEl.ValueKind == JsonValueKind.Array && hEl.GetArrayLength() > 0)
                        {
                            var v = hEl[0];
                            triggerHour = v.ValueKind == JsonValueKind.String ? (int.TryParse(v.GetString(), out var hi) ? hi : (int?)null)
                                       : v.ValueKind == JsonValueKind.Number ? v.GetInt32() : (int?)null;
                        }
                        if (schedEl.TryGetProperty("minutes", out var mnEl) && mnEl.ValueKind == JsonValueKind.Array && mnEl.GetArrayLength() > 0)
                        {
                            var v = mnEl[0];
                            triggerMinute = v.ValueKind == JsonValueKind.Number ? v.GetInt32()
                                          : v.ValueKind == JsonValueKind.String ? (int.TryParse(v.GetString(), out var mi) ? mi : (int?)null)
                                          : (int?)null;
                        }
                    }
                    break; // only first recurrence trigger
                }
            }

            TryReadReplyAction(defEl, out var internalReplyHtml, out var externalReplyHtml, out var externalAudience);

            return new CloudScheduleDefinitionResult(PowerAutomateOutcome.Success, message, flowDisplayName, triggerDays, triggerHour, triggerMinute, triggerTz,
                perDaySchedule: TryReadSidecar(defEl, out var sidecarTz, out var sidecarGen, out var sidecarTemplates),
                sidecarTimeZone: sidecarTz,
                sidecarGeneratedAt: sidecarGen,
                triggerStartTimeUtc: triggerStartTimeUtc,
                internalReplyHtml: internalReplyHtml,
                externalReplyHtml: externalReplyHtml,
                externalAudience: externalAudience,
                templates: sidecarTemplates);
        }
        catch (Exception ex)
        {
            return new CloudScheduleDefinitionResult(PowerAutomateOutcome.OtherError,
                $"Failed to parse PowerShell definition result: {ex.Message}", expectedFlowDisplayName, null, null, null, null);
        }
    }

    /// <summary>
    /// Decodes the OofManager sidecar metadata stamped by the generator
    /// under <c>definition.parameters._oofmgr_source.defaultValue</c>.
    /// Returns null when the per-day schedule is missing (older flow or
    /// vacation flow) while still populating <paramref name="templates"/>
    /// when present.
    /// </summary>
    private static IReadOnlyDictionary<DayOfWeek, CloudDaySchedule>? TryReadSidecar(
        JsonElement defEl,
        out string? tz,
        out string? generatedAt,
        out IReadOnlyList<OofTemplate>? templates)
    {
        tz = null;
        generatedAt = null;
        templates = null;
        if (!defEl.TryGetProperty("parameters", out var paramsEl) || paramsEl.ValueKind != JsonValueKind.Object)
            return null;
        if (!paramsEl.TryGetProperty("_oofmgr_source", out var srcEl) || srcEl.ValueKind != JsonValueKind.Object)
            return null;
        if (!srcEl.TryGetProperty("defaultValue", out var dvEl) || dvEl.ValueKind != JsonValueKind.Object)
            return null;

        if (dvEl.TryGetProperty("tz", out var tzEl) && tzEl.ValueKind == JsonValueKind.String) tz = tzEl.GetString();
        if (dvEl.TryGetProperty("generatedAt", out var gEl) && gEl.ValueKind == JsonValueKind.String) generatedAt = gEl.GetString();

        if (dvEl.TryGetProperty("templates", out var tplsEl) && tplsEl.ValueKind == JsonValueKind.Array)
        {
            var parsedTemplates = new List<OofTemplate>();
            foreach (var tEl in tplsEl.EnumerateArray())
            {
                if (tEl.ValueKind != JsonValueKind.Object) continue;
                var name = tEl.TryGetProperty("name", out var nEl) && nEl.ValueKind == JsonValueKind.String ? nEl.GetString() : null;
                if (string.IsNullOrWhiteSpace(name)) continue;
                var intReply = tEl.TryGetProperty("internalReply", out var irEl) && irEl.ValueKind == JsonValueKind.String ? irEl.GetString() ?? string.Empty : string.Empty;
                var extReply = tEl.TryGetProperty("externalReply", out var erEl) && erEl.ValueKind == JsonValueKind.String ? erEl.GetString() ?? string.Empty : string.Empty;
                var audAll = !tEl.TryGetProperty("externalAudienceAll", out var aaEl) || aaEl.ValueKind != JsonValueKind.False;
                var updatedAt = DateTime.UtcNow;
                if (tEl.TryGetProperty("updatedAt", out var upEl) && upEl.ValueKind == JsonValueKind.String &&
                    DateTimeOffset.TryParse(upEl.GetString(), System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out var parsedUp))
                {
                    updatedAt = parsedUp.UtcDateTime;
                }
                parsedTemplates.Add(new OofTemplate
                {
                    Name = name!.Trim(),
                    InternalReply = intReply,
                    ExternalReply = extReply,
                    ExternalAudienceAll = audAll,
                    CreatedAt = updatedAt,
                    UpdatedAt = updatedAt,
                });
            }
            templates = parsedTemplates;
        }

        if (!dvEl.TryGetProperty("days", out var daysEl) || daysEl.ValueKind != JsonValueKind.Object)
            return null;

        var map = new Dictionary<DayOfWeek, CloudDaySchedule>();
        foreach (var dayProp in daysEl.EnumerateObject())
        {
            if (!Enum.TryParse<DayOfWeek>(dayProp.Name, true, out var dow)) continue;
            if (dayProp.Value.ValueKind != JsonValueKind.Object) continue;
            var dEl = dayProp.Value;
            var isWorkday = dEl.TryGetProperty("workday", out var wEl) && wEl.ValueKind == JsonValueKind.True;
            if (!isWorkday && wEl.ValueKind == JsonValueKind.False) isWorkday = false;
            var start = ParseHm(dEl.TryGetProperty("start", out var sEl) ? sEl.GetString() : null);
            var end = ParseHm(dEl.TryGetProperty("end", out var eEl) ? eEl.GetString() : null);
            map[dow] = new CloudDaySchedule(isWorkday, start, end);
        }
        return map.Count == 0 ? null : map;

        static TimeSpan ParseHm(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return TimeSpan.Zero;
            var parts = s!.Split(':');
            if (parts.Length < 2) return TimeSpan.Zero;
            return int.TryParse(parts[0], out var h) && int.TryParse(parts[1], out var m)
                ? new TimeSpan(h, m, 0) : TimeSpan.Zero;
        }
    }

    // Pulls the reply text + audience out of the "Set up automatic replies (V2)"
    // action's inputs.parameters. Field names match what the generator writes
    // (body/automaticRepliesSetting/internalReplyMessage etc.). Any action key
    // containing "automatic replies" matches, so the lookup survives a future
    // localized rename of the action label in the designer.
    private static void TryReadReplyAction(JsonElement defEl, out string? internalReplyHtml, out string? externalReplyHtml, out string? externalAudience)
    {
        internalReplyHtml = null;
        externalReplyHtml = null;
        externalAudience = null;
        if (!defEl.TryGetProperty("actions", out var actionsEl) || actionsEl.ValueKind != JsonValueKind.Object)
            return;
        foreach (var actionProp in actionsEl.EnumerateObject())
        {
            if (actionProp.Value.ValueKind != JsonValueKind.Object) continue;
            if (!actionProp.Value.TryGetProperty("inputs", out var inputsEl) || inputsEl.ValueKind != JsonValueKind.Object) continue;
            if (!inputsEl.TryGetProperty("parameters", out var paramsEl) || paramsEl.ValueKind != JsonValueKind.Object) continue;
            // Match by parameter presence rather than action key so a future
            // rename (e.g. localization) doesn't break the compare.
            if (!paramsEl.TryGetProperty("body/automaticRepliesSetting/internalReplyMessage", out var iEl)) continue;
            if (iEl.ValueKind == JsonValueKind.String) internalReplyHtml = iEl.GetString();
            if (paramsEl.TryGetProperty("body/automaticRepliesSetting/externalReplyMessage", out var eEl) && eEl.ValueKind == JsonValueKind.String)
                externalReplyHtml = eEl.GetString();
            if (paramsEl.TryGetProperty("body/automaticRepliesSetting/externalAudience", out var aEl) && aEl.ValueKind == JsonValueKind.String)
                externalAudience = aEl.GetString();
            return;
        }
    }
    public async Task<CloudScheduleImportResult> ImportCloudScheduleSolutionAsync(
        string solutionZipPath,
        string solutionUniqueName,
        Guid workflowId,
        string expectedFlowDisplayName,
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        IProgress<string>? progress = null,
        CancellationToken ct = default,
        IReadOnlyList<PowerAutomateConnectionReference>? connectionReferences = null,
        IReadOnlyList<string>? expectedFlowDisplayNames = null)
    {
        if (!File.Exists(solutionZipPath))
        {
            return new CloudScheduleImportResult(
                CloudScheduleImportOutcome.OtherError,
                $"Solution zip not found at '{solutionZipPath}'.",
                null, null, null, null);
        }

        var targetFlowNames = expectedFlowDisplayNames is { Count: > 0 }
            ? expectedFlowDisplayNames.Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
            : (!string.IsNullOrWhiteSpace(expectedFlowDisplayName) ? new[] { expectedFlowDisplayName } : Array.Empty<string>());

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
                ct,
                connectionReferences,
                targetFlowNames).ConfigureAwait(false);

            if (cachedResult.Outcome == CloudScheduleImportOutcome.Success)
            {
                SaveImportEnvironmentCache(upnHint, cachedResult);
                SaveConnectionCache(upnHint, cachedResult.EnvironmentId ?? cachedEnvironment.EnvironmentId, cachedResult.BoundConnections);
                var verifiedRefs = cachedResult.VerifiedFlows.Where(i => i.FlowReference != null).Select(i => i.FlowReference!).ToArray();
                if (verifiedRefs.Length > 0) SaveFlowReferenceCache(upnHint, string.Empty, verifiedRefs);
                else SaveImportedFlowCandidateCache(upnHint, expectedFlowDisplayName, cachedResult);
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
            ct,
            connectionReferences,
            targetFlowNames).ConfigureAwait(false);

        if (result.Outcome == CloudScheduleImportOutcome.Success)
        {
            SaveImportEnvironmentCache(upnHint, result);
            SaveConnectionCache(upnHint, result.EnvironmentId, result.BoundConnections);
            var verifiedRefs = result.VerifiedFlows.Where(i => i.FlowReference != null).Select(i => i.FlowReference!).ToArray();
            if (verifiedRefs.Length > 0) SaveFlowReferenceCache(upnHint, string.Empty, verifiedRefs);
            else SaveImportedFlowCandidateCache(upnHint, expectedFlowDisplayName, result);
        }

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
        CancellationToken ct,
        IReadOnlyList<PowerAutomateConnectionReference>? connectionReferences = null,
        IReadOnlyList<string>? expectedFlowDisplayNames = null)
    {
        // Serialize the solution's connection references so the import script
        // can try to auto-bind each to an existing connection of the same
        // connector (via a pac deployment settings file). Empty/null when the
        // caller doesn't know them — the script then imports without binding.
        var connRefsJson = string.Empty;
        if (connectionReferences != null && connectionReferences.Count > 0)
        {
            var payload = connectionReferences
                .Where(c => !string.IsNullOrWhiteSpace(c.LogicalName) && !string.IsNullOrWhiteSpace(c.ConnectorId))
                .Select(c => new { logical = c.LogicalName, connectorId = c.ConnectorId })
                .ToList();
            if (payload.Count > 0)
                connRefsJson = JsonSerializer.Serialize(payload);
        }

        var cachedConnsJson = GetCachedConnectionsJson(upnHint, cachedEnvironment?.EnvironmentId);
        var flowNamesJson = expectedFlowDisplayNames is { Count: > 0 } ? JsonSerializer.Serialize(expectedFlowDisplayNames) : string.Empty;
        var cachedFlowMapJson = string.Empty;
        if (!string.IsNullOrWhiteSpace(cachedEnvironment?.EnvironmentId) && expectedFlowDisplayNames is { Count: > 0 })
        {
            var flowMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var dn in expectedFlowDisplayNames)
            {
                var r = GetCachedFlowReference(upnHint, dn, cachedEnvironment!.EnvironmentId);
                if (r != null && !string.IsNullOrWhiteSpace(r.FlowName))
                    flowMap[dn] = r.FlowName;
            }
            if (flowMap.Count == expectedFlowDisplayNames.Count)
                cachedFlowMapJson = JsonSerializer.Serialize(flowMap);
        }
        var envVars = new Dictionary<string, string>
        {
            ["OOFMGR_PA_UPN"]          = upnHint ?? string.Empty,
            ["OOFMGR_PA_DISPLAYNAME"]  = displayNameHint ?? string.Empty,
            ["OOFMGR_PA_ZIPPATH"]      = solutionZipPath,
            ["OOFMGR_PA_SOLNAME"]      = solutionUniqueName,
            ["OOFMGR_PA_WORKFLOWID"]   = workflowId.ToString("D"),
            ["OOFMGR_PA_FORCE"]        = forceOverwrite ? "1" : "0",
            ["OOFMGR_PA_CONNREFS"]     = connRefsJson,
            ["OOFMGR_PA_PRECREATE_CONN"] = connRefsJson.Length > 0 ? "1" : "0",
            ["OOFMGR_PA_CACHED_INSTANCEURL"] = cachedEnvironment?.InstanceUrl ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVID"]       = cachedEnvironment?.EnvironmentId ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVDISPLAY"]  = cachedEnvironment?.EnvironmentDisplayName ?? string.Empty,
            ["OOFMGR_PA_CACHED_CONNS"]       = cachedConnsJson,
            ["OOFMGR_PA_FLOWDISPLAYNAMES"]   = flowNamesJson,
            ["OOFMGR_PA_CACHED_FLOWMAP"]     = cachedFlowMapJson,
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
        _prefs.Set(ConnectionCacheV1Key, (string?)null);
        ClearFlowReferenceCache();
    }

    private string GetCachedConnectionsJson(string? upnHint, string? environmentId)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(environmentId)) return string.Empty;

        var raw = _prefs.GetString(ConnectionCacheV1Key);
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

        try
        {
            var cache = JsonSerializer.Deserialize<ConnectionCacheV1>(raw!);
            if (cache is null ||
                !string.Equals(cache.Upn, normalizedUpn, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(cache.EnvironmentId, environmentId!.Trim(), StringComparison.OrdinalIgnoreCase) ||
                cache.Connections is null ||
                cache.Connections.Count == 0)
            {
                return string.Empty;
            }
            return JsonSerializer.Serialize(cache.Connections);
        }
        catch
        {
            return string.Empty;
        }
    }

    private void SaveConnectionCache(string? upnHint, string? environmentId, IReadOnlyDictionary<string, string>? boundConnections)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(environmentId) || boundConnections is null || boundConnections.Count == 0)
            return;

        var trimmedEnv = environmentId!.Trim();
        ConnectionCacheV1? cache = null;
        var raw = _prefs.GetString(ConnectionCacheV1Key);
        if (!string.IsNullOrWhiteSpace(raw))
        {
            try { cache = JsonSerializer.Deserialize<ConnectionCacheV1>(raw!); }
            catch { cache = null; }
        }

        if (cache is null ||
            !string.Equals(cache.Upn, normalizedUpn, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(cache.EnvironmentId, trimmedEnv, StringComparison.OrdinalIgnoreCase))
        {
            cache = new ConnectionCacheV1 { Upn = normalizedUpn, EnvironmentId = trimmedEnv };
        }

        foreach (var kv in boundConnections)
        {
            if (!string.IsNullOrWhiteSpace(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
                cache.Connections[kv.Key.Trim()] = kv.Value.Trim();
        }

        if (cache.Connections.Count > 0)
            _prefs.Set(ConnectionCacheV1Key, JsonSerializer.Serialize(cache));
    }

    private sealed class ConnectionCacheV1
    {
        public string Upn { get; set; } = string.Empty;
        public string EnvironmentId { get; set; } = string.Empty;
        public Dictionary<string, string> Connections { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private CachedFlowReference? GetCachedFlowReference(string? upnHint, string expectedFlowDisplayName, string? expectedEnvironmentId)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(expectedFlowDisplayName)) return null;

        lock (_flowCacheGate)
        {
            var cache = LoadFlowReferenceCache(normalizedUpn);
            var entry = cache.Entries.FirstOrDefault(e =>
                string.Equals(e.DisplayName, expectedFlowDisplayName, StringComparison.OrdinalIgnoreCase));
            if (entry is null || string.IsNullOrWhiteSpace(entry.EnvironmentName) || string.IsNullOrWhiteSpace(entry.FlowName)) return null;
            if (!string.IsNullOrWhiteSpace(expectedEnvironmentId) &&
                !string.Equals(entry.EnvironmentName, expectedEnvironmentId, StringComparison.OrdinalIgnoreCase)) return null;
            return new CachedFlowReference(entry.EnvironmentName, entry.FlowName, entry.DisplayName);
        }
    }

    private void SaveFlowReferenceCache(string? upnHint, string expectedFlowDisplayName, IReadOnlyList<PowerAutomateFlowReference> flowReferences)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null) return;

        var valid = flowReferences.Where(f =>
            !string.IsNullOrWhiteSpace(f.EnvironmentName) &&
            !string.IsNullOrWhiteSpace(f.FlowName) &&
            !string.IsNullOrWhiteSpace(f.DisplayName)).ToArray();
        if (valid.Length == 0) return;

        lock (_flowCacheGate)
        {
            var cache = LoadFlowReferenceCache(normalizedUpn);
            foreach (var reference in valid)
                UpsertFlowReference(cache, reference.DisplayName, reference.EnvironmentName, reference.FlowName);
            SaveFlowReferenceCache(cache);
        }
    }

    private void SaveImportedFlowCandidateCache(string? upnHint, string expectedFlowDisplayName, CloudScheduleImportResult result)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(expectedFlowDisplayName) ||
            string.IsNullOrWhiteSpace(result.EnvironmentId) || string.IsNullOrWhiteSpace(result.WorkflowId)) return;

        lock (_flowCacheGate)
        {
            var cache = LoadFlowReferenceCache(normalizedUpn);
            if (!cache.Entries.Any(e => string.Equals(e.DisplayName, expectedFlowDisplayName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(e.EnvironmentName, result.EnvironmentId, StringComparison.OrdinalIgnoreCase)))
            {
                UpsertFlowReference(cache, expectedFlowDisplayName, result.EnvironmentId!, result.WorkflowId!);
                SaveFlowReferenceCache(cache);
            }
        }
    }

    private void RemoveCachedFlowReference(string? upnHint, string expectedFlowDisplayName)
    {
        var normalizedUpn = NormalizeCacheUpn(upnHint);
        if (normalizedUpn is null || string.IsNullOrWhiteSpace(expectedFlowDisplayName)) return;
        lock (_flowCacheGate)
        {
            var cache = LoadFlowReferenceCache(normalizedUpn);
            if (cache.Entries.RemoveAll(e => string.Equals(e.DisplayName, expectedFlowDisplayName, StringComparison.OrdinalIgnoreCase)) > 0)
                SaveFlowReferenceCache(cache);
        }
    }

    private FlowReferenceCacheV2 LoadFlowReferenceCache(string normalizedUpn)
    {
        FlowReferenceCacheV2? cache = null;
        var json = _prefs.GetString(FlowReferencesCacheV2Key);
        if (!string.IsNullOrWhiteSpace(json))
        {
            try { cache = JsonSerializer.Deserialize<FlowReferenceCacheV2>(json!); }
            catch { _prefs.Set(FlowReferencesCacheV2Key, (string?)null); }
        }
        if (cache is null || !string.Equals(cache.Upn, normalizedUpn, StringComparison.OrdinalIgnoreCase))
            cache = new FlowReferenceCacheV2 { Upn = normalizedUpn };

        // One-time migration of the legacy singleton entry for this account.
        var legacyUpn = _prefs.GetString(FlowCacheUpnKey);
        var legacyEnv = _prefs.GetString(FlowCacheEnvironmentIdKey);
        var legacyDisplay = _prefs.GetString(FlowCacheDisplayNameKey);
        var legacyName = _prefs.GetString(FlowCacheNameKey);
        if (string.Equals(legacyUpn, normalizedUpn, StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(legacyEnv) && !string.IsNullOrWhiteSpace(legacyDisplay) && !string.IsNullOrWhiteSpace(legacyName))
        {
            UpsertFlowReference(cache, legacyDisplay!, legacyEnv!, legacyName!);
            SaveFlowReferenceCache(cache);
        }
        ClearLegacyFlowReferenceCache();
        return cache;
    }

    private void SaveFlowReferenceCache(FlowReferenceCacheV2 cache)
    {
        _prefs.Set(FlowReferencesCacheV2Key, JsonSerializer.Serialize(cache));
    }

    private static void UpsertFlowReference(FlowReferenceCacheV2 cache, string displayName, string environmentName, string flowName)
    {
        cache.Entries.RemoveAll(e => string.Equals(e.DisplayName, displayName, StringComparison.OrdinalIgnoreCase));
        cache.Entries.Add(new FlowReferenceCacheEntry
        {
            DisplayName = displayName.Trim(),
            EnvironmentName = environmentName.Trim(),
            FlowName = flowName.Trim(),
        });
    }

    private void ClearFlowReferenceCache()
    {
        lock (_flowCacheGate)
        {
            using var batch = _prefs.BeginBatch();
            _prefs.Set(FlowReferencesCacheV2Key, (string?)null);
            ClearLegacyFlowReferenceCache();
        }
    }

    private void ClearLegacyFlowReferenceCache()
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

    private sealed class FlowReferenceCacheV2
    {
        public string Upn { get; set; } = string.Empty;
        public List<FlowReferenceCacheEntry> Entries { get; set; } = new();
    }

    private sealed class FlowReferenceCacheEntry
    {
        public string DisplayName { get; set; } = string.Empty;
        public string EnvironmentName { get; set; } = string.Empty;
        public string FlowName { get; set; } = string.Empty;
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
        // Status is a read-only probe that runs automatically at startup (and
        // after vacation operations). It must never pop a sign-in dialog on its
        // own, so run the child window hidden — no visible console flash at
        // launch. If silent auth isn't possible the probe degrades to an
        // "Unknown" banner and the user authenticates via an explicit action
        // (import / toggle / compare), which still uses the visible console.
        var json = await RunFlowScriptAsync("Get-Status", upnHint, displayNameHint, expectedFlowDisplayName, ct, progress, interactive: false).ConfigureAwait(false);
        var result = ParseStatusResult(json.Json, json.ExitCode);
        if (result.Outcome == PowerAutomateOutcome.Success)
            SaveFlowReferenceCache(upnHint, expectedFlowDisplayName, result.FlowReferences);
        else if (result.Outcome == PowerAutomateOutcome.NoFlowFound)
            RemoveCachedFlowReference(upnHint, expectedFlowDisplayName);
        return result;
    }

    private async Task<PowerAutomateResult> RunOperationAsync(string verb, string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress, CancellationToken ct)
    {
        var json = await RunFlowScriptAsync(verb, upnHint, displayNameHint, expectedFlowDisplayName, ct, progress).ConfigureAwait(false);
        var result = ParseResult(json.Json, json.ExitCode);
        if (result.Outcome == PowerAutomateOutcome.Success)
            SaveFlowReferenceCache(upnHint, expectedFlowDisplayName, result.FlowReferences);
        else if (result.Outcome == PowerAutomateOutcome.NoFlowFound)
            RemoveCachedFlowReference(upnHint, expectedFlowDisplayName);
        return result;
    }

    private async Task<PowerAutomateFlowBatchResult> RunFlowsBatchVerbAsync(string verb, string? upnHint, string? displayNameHint, IReadOnlyList<string> expectedFlowDisplayNames, IProgress<string>? progress, CancellationToken ct, bool interactive)
    {
        var names = expectedFlowDisplayNames.Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (names.Length == 0)
            return new PowerAutomateFlowBatchResult(PowerAutomateOutcome.OtherError, "No flow names were supplied.", Array.Empty<PowerAutomateFlowBatchItem>());

        var child = await RunFlowScriptAsync(verb, upnHint, displayNameHint, names[0], ct, progress, interactive: interactive, expectedFlowDisplayNames: names).ConfigureAwait(false);
        var result = ParseBatchResult(child.Json, child.ExitCode, names);
        var references = result.Items.Where(i => i.FlowReference != null).Select(i => i.FlowReference!).ToArray();
        if (references.Length > 0) SaveFlowReferenceCache(upnHint, string.Empty, references);
        foreach (var item in result.Items.Where(i => i.Outcome == PowerAutomateOutcome.NoFlowFound))
            RemoveCachedFlowReference(upnHint, item.DisplayName);
        return result;
    }

    private static PowerAutomateFlowBatchResult ParseBatchResult(string? resultJson, int exitCode, IReadOnlyList<string> expectedNames)
    {
        if (exitCode == -1 || string.IsNullOrWhiteSpace(resultJson))
            return new PowerAutomateFlowBatchResult(PowerAutomateOutcome.SignInFailed, "Power Automate batch operation did not complete.",
                expectedNames.Select(n => new PowerAutomateFlowBatchItem(n, PowerAutomateOutcome.SignInFailed, PowerAutomateFlowState.Unknown, "Operation did not complete.")).ToArray());
        try
        {
            using var doc = JsonDocument.Parse(resultJson!);
            var root = doc.RootElement;
            var rootOutcomeText = root.TryGetProperty("Outcome", out var ro) ? ro.GetString() : null;
            var overallMessage = root.TryGetProperty("Message", out var msg) ? msg.GetString() ?? "" : "";
            var rootOutcome = rootOutcomeText switch
            {
                "Success" => PowerAutomateOutcome.Success,
                "NoFlowFound" => PowerAutomateOutcome.NoFlowFound,
                "SignInFailed" => PowerAutomateOutcome.SignInFailed,
                "SolutionAwareBlocked" => PowerAutomateOutcome.SolutionAwareBlocked,
                _ => PowerAutomateOutcome.OtherError,
            };

            var items = new List<PowerAutomateFlowBatchItem>();
            if (root.TryGetProperty("BatchItems", out var array) && array.ValueKind == JsonValueKind.Array)
            {
                foreach (var el in array.EnumerateArray())
                {
                    var name = el.TryGetProperty("DisplayName", out var n) ? n.GetString() ?? "" : "";
                    var outcomeText = el.TryGetProperty("Outcome", out var o) ? o.GetString() : null;
                    var stateText = el.TryGetProperty("State", out var s) ? s.GetString() : null;
                    var message = el.TryGetProperty("Message", out var m) ? m.GetString() ?? "" : "";
                    PowerAutomateFlowReference? reference = null;
                    if (el.TryGetProperty("FlowReference", out var r) && r.ValueKind == JsonValueKind.Object)
                    {
                        var env = r.TryGetProperty("EnvironmentName", out var re) ? re.GetString() : null;
                        var flow = r.TryGetProperty("FlowName", out var rf) ? rf.GetString() : null;
                        var display = r.TryGetProperty("DisplayName", out var rd) ? rd.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(env) && !string.IsNullOrWhiteSpace(flow) && !string.IsNullOrWhiteSpace(display))
                            reference = new PowerAutomateFlowReference(env!, flow!, display!);
                    }
                    var outcome = outcomeText switch { "Success" => PowerAutomateOutcome.Success, "NoFlowFound" => PowerAutomateOutcome.NoFlowFound, "SignInFailed" => PowerAutomateOutcome.SignInFailed, _ => PowerAutomateOutcome.OtherError };
                    var state = stateText switch { "On" => PowerAutomateFlowState.On, "Off" => PowerAutomateFlowState.Off, "NotFound" => PowerAutomateFlowState.NotFound, _ => PowerAutomateFlowState.Unknown };
                    items.Add(new PowerAutomateFlowBatchItem(name, outcome, state, message, reference));
                }
            }

            if (items.Count == 0)
            {
                var fallbackState = rootOutcome == PowerAutomateOutcome.NoFlowFound ? PowerAutomateFlowState.NotFound : PowerAutomateFlowState.Unknown;
                foreach (var expectedName in expectedNames)
                    items.Add(new PowerAutomateFlowBatchItem(expectedName, rootOutcome, fallbackState, overallMessage));
                return new PowerAutomateFlowBatchResult(rootOutcome, overallMessage, items);
            }

            return new PowerAutomateFlowBatchResult(rootOutcome, overallMessage, items);
        }
        catch (Exception ex)
        {
            return new PowerAutomateFlowBatchResult(PowerAutomateOutcome.OtherError, $"Failed to parse batch result: {ex.Message}", Array.Empty<PowerAutomateFlowBatchItem>());
        }
    }

    private Task<ChildResult> RunFlowScriptAsync(string verb, string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct, IProgress<string>? progress = null, bool interactive = true, IReadOnlyList<string>? expectedFlowDisplayNames = null)
    {
        var cachedEnvironment = GetCachedImportEnvironment(upnHint);
        var cachedFlowReference = GetCachedFlowReference(upnHint, expectedFlowDisplayName, cachedEnvironment?.EnvironmentId);
        if (cachedFlowReference == null && expectedFlowDisplayNames != null)
        {
            foreach (var candidateName in expectedFlowDisplayNames)
            {
                cachedFlowReference = GetCachedFlowReference(upnHint, candidateName, cachedEnvironment?.EnvironmentId);
                if (cachedFlowReference != null) break;
            }
        }
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
            ["OOFMGR_PA_FLOWDISPLAYNAMES"]  = expectedFlowDisplayNames is { Count: > 0 } ? JsonSerializer.Serialize(expectedFlowDisplayNames) : string.Empty,
            ["OOFMGR_PA_CACHED_ENVID"]      = cachedEnvironment?.EnvironmentId ?? cachedFlowReference?.EnvironmentName ?? string.Empty,
            ["OOFMGR_PA_CACHED_ENVDISPLAY"] = cachedEnvironment?.EnvironmentDisplayName ?? string.Empty,
            ["OOFMGR_PA_CACHED_FLOWNAME"]   = cachedFlowReference?.FlowName ?? string.Empty,
        };
        return RunPowerShellChildAsync(
            BuildScript(),
            envVars,
            timeout: TimeSpan.FromMinutes(5),
            ct: ct,
            progress: progress,
            interactive: interactive);
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
        IProgress<string>? progress = null,
        bool interactive = true)
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
                // Visible console only for interactive paths (import / toggle)
                // where ADAL's sign-in dialog must anchor to a real console.
                // Background read-only status probes pass interactive:false so
                // the cmd/powershell child runs hidden — no startup flash.
                CreateNoWindow = !interactive,
            };
            psi.EnvironmentVariables["OOFMGR_PA_MODULES"]    = modulesDir;
            psi.EnvironmentVariables["OOFMGR_APPDIR"]        = AppContext.BaseDirectory;
            psi.EnvironmentVariables["OOFMGR_PA_RESULTFILE"] = resultFile;
            psi.EnvironmentVariables["OOFMGR_PA_PROGRESSFILE"] = progressFile;
            foreach (var kv in envVars)
            {
                psi.EnvironmentVariables[kv.Key] = kv.Value;
            }

            using var proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
            var exitTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            proc.Exited += (_, _) => exitTcs.TrySetResult(true);
            proc.Start();
            if (proc.HasExited) exitTcs.TrySetResult(true);

            var deadline = DateTime.UtcNow + timeout;
            var progressReadOffset = 0L;

            using (ct.Register(() => { try { if (!proc.HasExited) proc.Kill(); } catch { } }))
            {
                while (!exitTcs.Task.IsCompleted && !proc.HasExited)
                {
                    ReportProgressLines(progressFile, ref progressReadOffset, progress);
                    var remaining = deadline - DateTime.UtcNow;
                    if (remaining <= TimeSpan.Zero)
                    {
                        try { proc.Kill(); } catch { }
                        return new ChildResult(null, -1);
                    }

                    var delay = remaining < TimeSpan.FromMilliseconds(250) ? remaining : TimeSpan.FromMilliseconds(250);
                    await Task.WhenAny(exitTcs.Task, Task.Delay(delay, ct)).ConfigureAwait(false);
                    ct.ThrowIfCancellationRequested();
                }
                proc.WaitForExit();
                ReportProgressLines(progressFile, ref progressReadOffset, progress);
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

    private static void ReportProgressLines(string progressFile, ref long readOffset, IProgress<string>? progress)
    {
        if (progress is null || !File.Exists(progressFile)) return;

        try
        {
            using var fs = new FileStream(progressFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (fs.Length <= readOffset) return;
            fs.Seek(readOffset, SeekOrigin.Begin);
            using var reader = new StreamReader(fs, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            var chunk = reader.ReadToEnd();
            if (string.IsNullOrEmpty(chunk)) return;

            var lastNewline = chunk.LastIndexOf('\n');
            if (lastNewline < 0) return;

            var completeSlice = chunk.Substring(0, lastNewline + 1);
            readOffset += Encoding.UTF8.GetByteCount(completeSlice);

            var lines = completeSlice.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim();
                if (!string.IsNullOrWhiteSpace(line)) progress.Report(line);
            }
        }
        catch { }
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
            Dictionary<string, string>? boundConns = null;
            if (root.TryGetProperty("Connections", out var connsEl) && connsEl.ValueKind == JsonValueKind.Object)
            {
                boundConns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in connsEl.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(prop.Value.GetString()))
                        boundConns[prop.Name] = prop.Value.GetString()!;
                }
            }

            List<PowerAutomateFlowBatchItem>? verifiedFlows = null;
            if (root.TryGetProperty("BatchItems", out var batchEl) && batchEl.ValueKind == JsonValueKind.Array)
            {
                verifiedFlows = new List<PowerAutomateFlowBatchItem>();
                foreach (var el in batchEl.EnumerateArray())
                {
                    var name = el.TryGetProperty("DisplayName", out var n) ? n.GetString() ?? "" : "";
                    var outcomeText = el.TryGetProperty("Outcome", out var o) ? o.GetString() : null;
                    var stateText = el.TryGetProperty("State", out var s) ? s.GetString() : null;
                    var message = el.TryGetProperty("Message", out var m) ? m.GetString() ?? "" : "";
                    PowerAutomateFlowReference? reference = null;
                    if (el.TryGetProperty("FlowReference", out var r) && r.ValueKind == JsonValueKind.Object)
                    {
                        var env = r.TryGetProperty("EnvironmentName", out var re) ? re.GetString() : null;
                        var flow = r.TryGetProperty("FlowName", out var rf) ? rf.GetString() : null;
                        var display = r.TryGetProperty("DisplayName", out var rd) ? rd.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(env) && !string.IsNullOrWhiteSpace(flow) && !string.IsNullOrWhiteSpace(display))
                            reference = new PowerAutomateFlowReference(env!, flow!, display!);
                    }
                    var itemOutcome = outcomeText switch { "Success" => PowerAutomateOutcome.Success, "NoFlowFound" => PowerAutomateOutcome.NoFlowFound, "SignInFailed" => PowerAutomateOutcome.SignInFailed, _ => PowerAutomateOutcome.OtherError };
                    var itemState = stateText switch { "On" => PowerAutomateFlowState.On, "Off" => PowerAutomateFlowState.Off, "NotFound" => PowerAutomateFlowState.NotFound, _ => PowerAutomateFlowState.Unknown };
                    verifiedFlows.Add(new PowerAutomateFlowBatchItem(name, itemOutcome, itemState, message, reference));
                }
            }

            return new CloudScheduleImportResult(
                outcome,
                Read("Message") ?? string.Empty,
                Read("EnvironmentId"),
                Read("EnvironmentDisplayName"),
                Read("InstanceUrl"),
                Read("WorkflowId"),
                boundConns,
                verifiedFlows);
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
            var stateStr = root.TryGetProperty("State", out var sEl) ? sEl.GetString() : null;

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
                _ => PowerAutomateFlowState.Unknown,
            };
            return new PowerAutomateResult(outcome, message, flows, flowReferences, state);
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
$flowDisplayNamesJson = $env:OOFMGR_PA_FLOWDISPLAYNAMES
$flowDisplayNames = @()
if ($flowDisplayNamesJson) {
    try {
        $parsedFlowNames = ConvertFrom-Json -InputObject $flowDisplayNamesJson
        # Windows PowerShell 5.1 can preserve a JSON array as one nested
        # Object[] when wrapped directly in @(...). Flatten explicitly.
        foreach ($parsedFlowName in @($parsedFlowNames)) {
            foreach ($nameValue in @($parsedFlowName)) {
                if ($nameValue) { $flowDisplayNames += [string]$nameValue }
            }
        }
    } catch { Trace ""flow display-name JSON parse failed: $($_.Exception.Message)"" }
}
if ($flowDisplayNames.Count -eq 0 -and $flowDisplayName) { $flowDisplayNames = @($flowDisplayName) }
$flowDisplayNamesTrace = $flowDisplayNames -join ' | '
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
# script never returns we know exactly which step blocked. Per-pid file
# avoids collisions when multiple Power Automate cmds run concurrently
# (startup status refreshes overlap with user-triggered compare/toggle).
$debugLog = Join-Path $env:TEMP (""oofmgr-pa-debug-$PID.log"")
$debugWriter = $null
try {
    $debugFs = [System.IO.FileStream]::new($debugLog, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    $debugWriter = [System.IO.StreamWriter]::new($debugFs, [System.Text.UTF8Encoding]::new($false))
    $debugWriter.AutoFlush = $true
} catch {}
$progressWriter = $null
if ($progressFile) {
    try {
        $progFs = [System.IO.FileStream]::new($progressFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $progressWriter = [System.IO.StreamWriter]::new($progFs, [System.Text.UTF8Encoding]::new($false))
        $progressWriter.AutoFlush = $true
    } catch {}
}
function Trace($m) {
    if ($debugWriter) {
        try { $debugWriter.WriteLine(('{0}  {1}' -f (Get-Date -Format o), $m)) } catch {}
    }
}
function Report($m) {
    if (-not $m) { return }
    Trace (""progress: "" + $m)
    if ($progressWriter) {
        try { $progressWriter.WriteLine([string]$m) } catch {}
    }
}
Trace (""target flow display names: "" + $flowDisplayNamesTrace)
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

function Emit($outcome, $message, $flows, $state = $null, $flowReferences = @(), $definition = $null, $batchItems = @()) {
    $obj = [ordered]@{
        Outcome = $outcome
        Message = $message
        Flows   = @($flows)
        State   = $state
        FlowReferences = @($flowReferences)
        Definition = $definition
        BatchItems = @($batchItems)
    }
    $json = ($obj | ConvertTo-Json -Compress -Depth 25)
    if ($resultFile) {
        [System.IO.File]::WriteAllText($resultFile, $json, [System.Text.UTF8Encoding]::new($false))
    }
    Trace ""emit: outcome=$outcome message=$message""
    if ($progressWriter) { try { $progressWriter.Dispose() } catch {} }
    if ($debugWriter) { try { $debugWriter.Dispose() } catch {} }
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
        Report 'Signed in. Looking for your weekly schedule flow...'
    } catch {
        Trace ""Add-PowerAppsAccount FAILED: $($_.Exception.Message)""
        Emit 'SignInFailed' (""Sign-in to Power Automate failed: "" + $_.Exception.Message) @()
    }

    # Every flow operation is scoped to the signed-in OOF Manager user's own
    # Developer environment. The WPF caller supplies the Exchange-resolved
    # friendly display name (for example 'Sandy Sun'); we require an exact,
    # case-insensitive environment DisplayName match and never probe Default,
    # another user's Developer environment, or arbitrary tenant environments.
    # The exact flow display name remains the second ownership boundary inside
    # that one environment.
    $matched         = @()

    function Add-MatchedFlow($envName, $flow, $sourceLabel) {
        if (-not $flow) { return $false }
        $dn = [string]$flow.DisplayName
        $flowName = [string]$flow.FlowName
        if ($dn -and $flowName -and ($flowDisplayNames -icontains $dn)) {
            # Keep the original object so callers can reuse a reliable state
            # snapshot without immediately issuing the same Get-Flow again.
            $script:matched += [PSCustomObject]@{
                EnvironmentName = $envName
                FlowName        = $flowName
                DisplayName     = $dn
                Enabled         = $flow.Enabled
                SourceFlow      = $flow
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

    function Get-CurrentUserDeveloperEnvFast() {
        if (-not $displayName -or ([string]$displayName).IndexOf('@') -ge 0) {
            Trace ""cannot select Developer environment: friendly display name is unavailable ('$displayName')""
            return $null
        }

        $developerEnvs = @()
        try {
            # Ask only for Developer environments, then require an exact friendly-
            # name match. This avoids enumerating every environment visible in a
            # large tenant and prevents us from probing another user's flows.
            $developerEnvUri = ""https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/environments?`$expand=permissions&`$filter=properties/environmentSku%20eq%20%27Developer%27&api-version={apiVersion}""
            Report ""Finding your Developer environment '$displayName'...""
            Trace 'before exact developer-env query'
            $developerResult = InvokeApi -Method GET -Route $developerEnvUri -ErrorAction Stop
            $developerEnvs = @($developerResult.value | ForEach-Object { Convert-RawEnvironment $_ } | Where-Object { $_ -and (Test-DeveloperEnvironment $_) })
            Trace ""after exact developer-env query: count=$($developerEnvs.Count)""
        } catch {
            Trace ""exact developer-env query failed: $($_.Exception.Message)""
            return $null
        }

        $exactDeveloperEnvs = @(Get-ExactDisplayNameEnvs $developerEnvs)
        Trace ""exact-display-name Developer environments: count=$($exactDeveloperEnvs.Count) displayName='$displayName'""
        if ($exactDeveloperEnvs.Count -eq 0) { return $null }

        # Duplicate display names are unusual. Prefer an environment where the
        # current user has an admin permission; otherwise use the API's first
        # exact match. In either case exactly one environment is ever scanned.
        $adminExact = @($exactDeveloperEnvs | Where-Object { Test-AdminEnvironmentForCurrentUser $_ })
        $selected = if ($adminExact.Count -gt 0) { $adminExact[0] } else { $exactDeveloperEnvs[0] }
        Trace ""selected current-user Developer environment displayName='$($selected.DisplayName)' id='$($selected.EnvironmentName)' adminExact=$($adminExact.Count -gt 0)""
        return $selected
    }

    $isBatchVerb = ($verb -in @('Ensure-Flows-On', 'Get-Flows-Status', 'Disable-Flows-Batch', 'Enable-Flows-Batch'))

    # A cached reference is safe only when it belongs to the signed-in user's
    # exact Developer environment. Never probe a stale/differently named cache.
    $cacheMatchesCurrentUser = $cachedEnvId -and $cachedEnvDisplay -and $displayName -and
        (([string]$cachedEnvDisplay).Trim() -ieq ([string]$displayName).Trim())
    $currentUserEnvironmentChecked = $false
    $targetEnvironmentId = $null
    if ($cacheMatchesCurrentUser) {
        $cachedEnvId = ([string]$cachedEnvId).Trim()
        $targetEnvironmentId = $cachedEnvId
        $cachedLabel = ([string]$cachedEnvDisplay).Trim()
        $cachedFlowName = if ($cachedFlowName -and ([string]$cachedFlowName).Trim().Length -gt 0) { ([string]$cachedFlowName).Trim() } else { $null }
        if ($cachedFlowName -and (-not $isBatchVerb)) {
            Report ""Checking saved cloud flow in your Developer environment '$cachedLabel'...""
            Trace ""validated cached flow candidate displayName='$cachedLabel' envId='$cachedEnvId' flowName='$cachedFlowName'""
            $null = Try-CachedFlow $cachedEnvId 'current-user-cached-flow' $cachedFlowName
        }

        if ($matched.Count -eq 0) {
            Report ""Checking your saved Developer environment '$cachedLabel'...""
            Trace ""validated cached env candidate displayName='$cachedLabel' id='$cachedEnvId'""
            $currentUserEnvironmentChecked = $true
            Scan-Env $cachedEnvId 'current-user-cached-env'
        }
    } elseif ($cachedEnvId) {
        Trace ""ignoring cached environment outside current-user scope: cachedDisplay='$cachedEnvDisplay' currentDisplay='$displayName' id='$cachedEnvId'""
    }

    $currentUserDeveloperEnv = $null
    if ($matched.Count -eq 0 -and -not $currentUserEnvironmentChecked) {
        $currentUserDeveloperEnv = Get-CurrentUserDeveloperEnvFast
        if ($currentUserDeveloperEnv) {
            Report ""Checking your Developer environment '$($currentUserDeveloperEnv.DisplayName)'...""
            $currentUserEnvironmentChecked = $true
            $targetEnvironmentId = [string]$currentUserDeveloperEnv.EnvironmentName
            Scan-Env $currentUserDeveloperEnv.EnvironmentName 'current-user-developer-env'
        }
    }
    if ($matched.Count -gt 0) {
        Trace 'matched in current-user Developer environment; no other environments will be checked'
    }

    # Strict scope: all flow operations target only the Developer environment
    # whose display name exactly matches the signed-in OOF Manager user.
    if ($matched.Count -eq 0 -and (-not $isBatchVerb)) {
        if (-not $displayName -or ([string]$displayName).IndexOf('@') -ge 0) {
            Emit 'NoFlowFound' 'The signed-in user friendly display name is unavailable, so OOF Manager could not identify your Developer environment.' @()
        }
        if (-not $currentUserEnvironmentChecked) {
            Emit 'NoFlowFound' (""No Developer environment named '"" + ([string]$displayName).Trim() + ""' was found for the signed-in OOF Manager user. No other environments were checked."") @()
        }
        Emit 'NoFlowFound' (""No flow named '"" + $flowDisplayName + ""' was found in your Developer environment '"" + ([string]$displayName).Trim() + ""'. No other environments were checked."") @()
    }

    # Legacy broad-scan implementation retained as unreachable reference until
    # the strict current-user path has been exercised across supported tenants.
    if ($false) {

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

    function Read-SnapshotEnabled([object]$flow) {
        $result = [pscustomobject]@{ Known = $false; Enabled = $null }
        try {
            $source = $flow.SourceFlow
            if (-not $source) { return $result }
            $enabledProp = $source.PSObject.Properties['Enabled']
            $rawState = if ($source.PSObject.Properties['Internal'] -and $source.Internal -and $source.Internal.properties) { [string]$source.Internal.properties.state } else { $null }
            if ($enabledProp -ne $null -and $enabledProp.Value -is [bool]) {
                # A direct Get-Flow-by-id object is authoritative. Bulk objects
                # are also usable when their internal state agrees with Enabled.
                $stateAgrees = (-not $rawState) -or
                    ($rawState -ieq 'Started' -and [bool]$enabledProp.Value) -or
                    ($rawState -ieq 'Stopped' -and -not [bool]$enabledProp.Value)
                if ($stateAgrees) {
                    $result.Known = $true
                    $result.Enabled = [bool]$enabledProp.Value
                    Trace ""using fetched flow state: $($flow.DisplayName) enabled='$($result.Enabled)' rawState='$rawState'""
                }
            }
        } catch { Trace ""snapshot state read failed for $($flow.DisplayName): $($_.Exception.Message)"" }
        return $result
    }

    function Read-CurrentOrSnapshotEnabled([object]$flow) {
        $snapshot = Read-SnapshotEnabled $flow
        if ($snapshot.Known) { return $snapshot }
        return Read-CurrentEnabled $flow
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

    function Invoke-FlowsToggleParallel([object[]]$flowsToToggle, [bool]$desiredEnabled) {
        $errorsByName = @{}
        if (-not $flowsToToggle -or $flowsToToggle.Count -eq 0) { return $errorsByName }
        $actionSegment = if ($desiredEnabled) { 'start' } else { 'stop' }
        $targetCmd = if ($desiredEnabled) { 'Enable-Flow' } else { 'Disable-Flow' }
        $flowEndpoint = if ($global:currentSession -and $global:currentSession.flowEndpoint) { [string]$global:currentSession.flowEndpoint } else { 'api.flow.microsoft.com' }
        $token = $null
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $token = Get-JwtToken -Audience 'https://service.flow.microsoft.com/'
        } catch {
            Trace ""parallel toggle token lookup failed ($($_.Exception.Message)); falling back to sequential cmdlets""
        }

        if (-not $token) {
            foreach ($f in $flowsToToggle) {
                try {
                    if ($f.EnvironmentName) { & $targetCmd -EnvironmentName $f.EnvironmentName -FlowName $f.FlowName -ErrorAction Stop | Out-Null }
                    else { & $targetCmd -FlowName $f.FlowName -ErrorAction Stop | Out-Null }
                } catch {
                    $errorsByName[[string]$f.DisplayName] = $_.Exception.Message
                }
            }
            return $errorsByName
        }

        $authHeader = if ($token -like 'Bearer *') { [string]$token } else { 'Bearer ' + [string]$token }
        $pendingTasks = @()
        foreach ($f in $flowsToToggle) {
            $dn = [string]$f.DisplayName
            try {
                $envSegment = ResolveEnvironment -OverrideId ([string]$f.EnvironmentName)
                $uri = ""https://$flowEndpoint/providers/Microsoft.ProcessSimple/environments/$envSegment/flows/$($f.FlowName)/$actionSegment`?api-version=2016-11-01""
                Trace ""starting parallel POST $actionSegment for '$dn' flow=$($f.FlowName)""
                $req = [System.Net.HttpWebRequest]::Create($uri)
                $req.Method = 'POST'
                $req.ContentLength = 0
                $req.UserAgent = 'PowerShell cmdlets 1.0'
                $req.Headers.Add('Authorization', $authHeader)
                $pendingTasks += [pscustomobject]@{
                    DisplayName = $dn
                    Task        = $req.GetResponseAsync()
                }
            } catch {
                $errorsByName[$dn] = $_.Exception.Message
            }
        }

        foreach ($item in $pendingTasks) {
            $dn = $item.DisplayName
            try {
                $resp = $item.Task.GetAwaiter().GetResult()
                if ($resp) { try { $resp.Dispose() } catch {} }
                Trace ""finished parallel POST $actionSegment for '$dn'""
            } catch {
                $ex = $_.Exception
                if ($ex.InnerException) { $ex = $ex.InnerException }
                Trace ""parallel POST $actionSegment failed for '$dn': $($ex.Message)""
                $errorsByName[$dn] = $ex.Message
            }
        }
        return $errorsByName
    }

    if ($verb -in @('Ensure-Flows-On', 'Disable-Flows-Batch', 'Enable-Flows-Batch')) {
        $desiredEnabled = ($verb -ne 'Disable-Flows-Batch')
        $targetCmd = if ($desiredEnabled) { 'Enable-Flow' } else { 'Disable-Flow' }
        $targetStateLabel = if ($desiredEnabled) { 'On' } else { 'Off' }
        $targetLower = $targetStateLabel.ToLowerInvariant()
        $batchItemsByName = @{}
        $toToggle = @()

        if ($verb -eq 'Ensure-Flows-On') {
            $missingNames = @($flowDisplayNames | Where-Object { $name = $_; -not ($matched | Where-Object { $_.DisplayName -ieq $name }) })
            if ($missingNames.Count -gt 0 -and $currentUserEnvironmentChecked) {
                Report ""Vacation flows are still registering. Waiting 10 seconds before one retry...""
                Start-Sleep -Seconds 10
                $matched = @()
                Scan-Env $targetEnvironmentId 'batch-registration-retry'
            }
        }

        foreach ($expectedName in $flowDisplayNames) {
            $f = $matched | Where-Object { $_.DisplayName -ieq $expectedName } | Select-Object -First 1
            if (-not $f) {
                $notFoundMsg = if ($verb -eq 'Ensure-Flows-On') { 'Flow was not found after the registration retry.' } else { 'Flow was not found.' }
                $batchItemsByName[$expectedName] = [pscustomobject]@{ DisplayName=$expectedName; Outcome='NoFlowFound'; State='NotFound'; Message=$notFoundMsg; FlowReference=$null }
                continue
            }

            $initial = Read-CurrentOrSnapshotEnabled $f
            if ($initial.Known -and ($initial.Enabled -eq $desiredEnabled)) {
                $batchItemsByName[$expectedName] = [pscustomobject]@{ DisplayName=$expectedName; Outcome='Success'; State=$targetStateLabel; Message=""Power Automate flow is already $targetLower.""; FlowReference=(New-FlowReference $f) }
                continue
            }
            $toToggle += $f
        }

        if ($toToggle.Count -gt 0) {
            $namesLabel = (@($toToggle | ForEach-Object { ""'$($_.DisplayName)'"" }) -join ', ')
            Report ""Turning $namesLabel $targetLower...""
            $toggleErrors = Invoke-FlowsToggleParallel $toToggle $desiredEnabled

            # Verify all toggled flows in one environment scan first before
            # falling back to per-flow retries.
            for ($vi = 0; $vi -lt 3; $vi++) {
                $matched = @()
                if ($targetEnvironmentId) {
                    Scan-Env $targetEnvironmentId ""batch-post-toggle-verify[$vi]""
                }
                $allVerified = $true
                foreach ($tf in $toToggle) {
                    $dn = [string]$tf.DisplayName
                    if ($toggleErrors.ContainsKey($dn)) { continue }
                    $refreshed = $matched | Where-Object { $_.DisplayName -ieq $dn } | Select-Object -First 1
                    $check = if ($refreshed) { Read-SnapshotEnabled $refreshed } else { Read-CurrentEnabled $tf }
                    if (-not $check.Known) { $check = Read-CurrentEnabled $tf }
                    if (-not ($check.Known -and ($check.Enabled -eq $desiredEnabled))) {
                        $allVerified = $false
                    }
                }
                if ($allVerified) { break }
                if ($vi -lt 2) { Start-Sleep -Seconds 3 }
            }

            foreach ($tf in $toToggle) {
                $dn = [string]$tf.DisplayName
                if ($toggleErrors.ContainsKey($dn)) {
                    $batchItemsByName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='OtherError'; State='Unknown'; Message=$toggleErrors[$dn]; FlowReference=(New-FlowReference $tf) }
                    continue
                }
                $refreshed = $matched | Where-Object { $_.DisplayName -ieq $dn } | Select-Object -First 1
                $after = if ($refreshed) { Read-SnapshotEnabled $refreshed } else { Read-CurrentEnabled $tf }
                if (-not $after.Known) { $after = Read-CurrentEnabled $tf }

                if ($after.Known -and ($after.Enabled -eq $desiredEnabled)) {
                    $batchItemsByName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='Success'; State=$targetStateLabel; Message=""Power Automate flow was turned $targetLower.""; FlowReference=(New-FlowReference $tf) }
                } elseif (-not $after.Known) {
                    $batchItemsByName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='Success'; State='Unknown'; Message=""$targetCmd was accepted but the final state could not be read.""; FlowReference=(New-FlowReference $tf) }
                } else {
                    $actualState = if ($after.Enabled) { 'On' } else { 'Off' }
                    $batchItemsByName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='OtherError'; State=$actualState; Message=""$targetCmd was accepted but the flow is still $($actualState.ToLowerInvariant()) (a connection reference is probably unbound).""; FlowReference=(New-FlowReference $tf) }
                }
            }
        }

        $batchItems = @()
        foreach ($expectedName in $flowDisplayNames) {
            if ($batchItemsByName.ContainsKey($expectedName)) {
                $batchItems += $batchItemsByName[$expectedName]
            }
        }

        $allMatchedTarget = @($batchItems | Where-Object { $_.Outcome -ne 'Success' -or $_.State -ne $targetStateLabel }).Count -eq 0
        $batchOutcome = if ($allMatchedTarget) { 'Success' } else { 'OtherError' }
        $batchMessage = if ($allMatchedTarget) { ""All requested Power Automate flows are $targetLower."" } else { ""One or more requested Power Automate flows could not be confirmed $targetLower."" }
        Emit $batchOutcome $batchMessage @($batchItems | ForEach-Object { $_.DisplayName }) $null (Get-FlowReferences $toToggle) $null $batchItems
    }

    if ($verb -eq 'Get-Flows-Status') {
        $batchItems = @()
        foreach ($expectedName in $flowDisplayNames) {
            $f = $matched | Where-Object { $_.DisplayName -ieq $expectedName } | Select-Object -First 1
            if (-not $f) {
                $batchItems += [pscustomobject]@{ DisplayName=$expectedName; Outcome='NoFlowFound'; State='NotFound'; Message='Flow was not found.'; FlowReference=$null }
                continue
            }

            $stateInfo = Read-CurrentOrSnapshotEnabled $f
            if ($stateInfo.Known) {
                $st = if ($stateInfo.Enabled) { 'On' } else { 'Off' }
                $batchItems += [pscustomobject]@{ DisplayName=$expectedName; Outcome='Success'; State=$st; Message=('Power Automate flow is ' + $st.ToLowerInvariant() + '.'); FlowReference=(New-FlowReference $f) }
            } else {
                $batchItems += [pscustomobject]@{ DisplayName=$expectedName; Outcome='OtherError'; State='Unknown'; Message='Power Automate flow was found, but its current state could not be read.'; FlowReference=(New-FlowReference $f) }
            }
        }
        Emit 'Success' 'Batch flow status check complete.' @($batchItems | ForEach-Object { $_.DisplayName }) $null (Get-FlowReferences $matched) $null $batchItems
    }

    if ($verb -eq 'Get-Status') {
        foreach ($f in $matched) {
            Report ""Found weekly schedule flow '$($f.DisplayName)'. Reading current state...""
            $freshState = Read-CurrentOrSnapshotEnabled $f
            if ($freshState.Known) {
                $state = if ($freshState.Enabled) { 'On' } else { 'Off' }
                $lowerState = $state.ToLowerInvariant()
                Report ""Power Automate flow is $lowerState.""
                Emit 'Success' ('Power Automate flow is ' + $lowerState) @($f.DisplayName) $state (Get-FlowReferences @($f))
            }
        }

        Emit 'OtherError' 'Power Automate flow was found, but its current state could not be read.' @($matched | ForEach-Object { $_.DisplayName }) 'Unknown'
    }

    if ($verb -eq 'Get-Definition') {
        foreach ($f in $matched) {
            Report ""Found cloud flow '$($f.DisplayName)'. Reading definition...""
            $definition = $null
            $defSource = '<none>'
            try {
                # If $f was matched via Try-CachedFlow (Get-Flow -FlowName),
                # $f.SourceFlow already holds the full object with definition /
                # definitionSummary. Only issue a second Get-Flow when $f came
                # from a bulk environment scan that omits definition properties.
                $full = $f.SourceFlow
                $hasCachedDef = $full -and $full.PSObject.Properties['Internal'] -and $full.Internal -and $full.Internal.properties -and
                    (($full.Internal.properties.PSObject.Properties['definition'] -and $full.Internal.properties.definition) -or
                     ($full.Internal.properties.PSObject.Properties['definitionSummary'] -and $full.Internal.properties.definitionSummary))
                if ($hasCachedDef) {
                    Trace ""reusing cached full flow object for definition env=$($f.EnvironmentName) flow=$($f.FlowName)""
                } else {
                    try {
                        Trace ""before Get-Flow full env=$($f.EnvironmentName) flow=$($f.FlowName)""
                        $full = Get-Flow -EnvironmentName $f.EnvironmentName -FlowName $f.FlowName -ErrorAction Stop
                        Trace ""after  Get-Flow full""
                    } catch {
                        Trace ""Get-Flow full FAILED: $($_.Exception.Message)""
                    }
                }
                if ($full -and $full.PSObject.Properties['Internal'] -and $full.Internal -and $full.Internal.properties) {
                    $props = $full.Internal.properties
                    $propNames = @($props.PSObject.Properties | ForEach-Object { $_.Name })
                    Trace ""flow Internal.properties keys: $($propNames -join ',')""

                    if ($props.PSObject.Properties['definition'] -and $props.definition) {
                        $definition = $props.definition
                        $defSource = 'definition'
                    } elseif ($props.PSObject.Properties['definitionSummary'] -and $props.definitionSummary) {
                        $definition = $props.definitionSummary
                        $defSource = 'definitionSummary'
                    }
                }
            } catch {
                Trace ""definition extract failed: $($_.Exception.Message)""
            }
            Trace ""definition source: $defSource""
            if ($definition -ne $null) {
                Emit 'Success' (""Cloud flow definition retrieved from "" + $defSource) @($f.DisplayName) $null (Get-FlowReferences @($f)) $definition
            }
            Emit 'OtherError' 'Cloud flow was found, but neither definition nor definitionSummary could be read.' @($f.DisplayName) 'Unknown'
        }

        Emit 'NoFlowFound' 'No cloud flow with the expected display name was found.' @() $null @()
    }

    $changed = @()
    $alreadyDesired = @()
    $errors  = @()
    $finalStateKnown = $false
    $finalEnabled = $null
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
            $freshState = Read-CurrentOrSnapshotEnabled $f
            if ($freshState.Known -and $freshState.Enabled -eq $desiredEnabled) {
                Trace ""skip $verb; fresh state already desired: $($f.DisplayName)""
                Report ""Power Automate flow is already $targetState. Finishing...""
                $alreadyDesired += $f.DisplayName
                $finalStateKnown = $true
                $finalEnabled = $desiredEnabled
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

            # Double-insurance: the cmdlet ACCEPTING the change is not proof
            # the flow reached the desired state. When a connection reference
            # is unbound, Enable-Flow is accepted but the flow silently stays
            # Off. Re-read the live state and only count it as changed if it
            # actually flipped. Retry a couple times to absorb write-
            # propagation lag (a real success breaks out on the first read, so
            # only the genuine still-Off / unbound case pays the wait). An
            # unreadable state is treated as accepted so a transient read
            # failure can't turn a real success into a false alarm.
            $verified = $false
            $lastKnown = $false
            $lastEnabled = $null
            for ($vi = 0; $vi -lt 3; $vi++) {
                $after = Read-CurrentEnabled $f
                $lastKnown = $after.Known
                $lastEnabled = $after.Enabled
                if ($after.Known -and ($after.Enabled -eq $desiredEnabled)) { $verified = $true; break }
                if (-not $after.Known) { break }
                if ($vi -lt 2) { Start-Sleep -Seconds 3 }
            }
            if ($verified) {
                Trace ""post-$verb verify OK: $($f.DisplayName) is $targetState""
                Report 'Power Automate accepted the change. Finishing...'
                $changed += $f.DisplayName
                $finalStateKnown = $true
                $finalEnabled = $desiredEnabled
            } elseif (-not $lastKnown) {
                Trace ""post-$verb verify inconclusive (state unreadable); treating as accepted: $($f.DisplayName)""
                Report 'Power Automate accepted the change. Finishing...'
                $changed += $f.DisplayName
            } else {
                $finalStateKnown = $true
                $finalEnabled = [bool]$lastEnabled
                $actual = if ($lastEnabled) { 'on' } else { 'off' }
                Trace ""post-$verb verify FAILED: $($f.DisplayName) still $actual, expected $targetState (connection reference likely unbound)""
                $errors += ($f.DisplayName + "": the change was accepted but the flow is still "" + $actual + "" (a connection reference is probably unbound)."")
            }
        } catch {
            Trace ""$verb FAILED on $($f.DisplayName): $($_.Exception.Message)""
            $errors += ($f.DisplayName + "": "" + $_.Exception.Message)
        }
    }

    $resultState = if ($finalStateKnown) { if ($finalEnabled) { 'On' } else { 'Off' } } else { $null }
    if ($changed.Count -eq 0 -and $alreadyDesired.Count -eq 0 -and $errors.Count -gt 0) {
        $msg = $errors -join '; '
        if ($msg -match 'solution' -or $msg -match 'managed' -or $msg -match 'not.*allowed') {
            Emit 'SolutionAwareBlocked' $msg @() $resultState (Get-FlowReferences $matched)
        }
        Emit 'OtherError' $msg @() $resultState (Get-FlowReferences $matched)
    }

    $action = if ($verb -eq 'Disable-Flow') { 'Turned off' } else { 'Turned on' }
    $state = if ($verb -eq 'Disable-Flow') { 'off' } else { 'on' }
    $msg = if ($changed.Count -gt 0) { ""$action Power Automate flow"" } else { ""Power Automate flow is already $state"" }
    if ($errors.Count -gt 0) { $msg = $msg + "" (with "" + $errors.Count + "" partial failure(s): "" + ($errors -join '; ') + "")"" }
    Report 'Done.'
    Emit 'Success' $msg @($changed + $alreadyDesired) $resultState (Get-FlowReferences $matched)

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
$appDir       = $env:OOFMGR_APPDIR
$upn          = $env:OOFMGR_PA_UPN
$displayName  = $env:OOFMGR_PA_DISPLAYNAME
$resultFile   = $env:OOFMGR_PA_RESULTFILE
$progressFile = $env:OOFMGR_PA_PROGRESSFILE
$zipPath      = $env:OOFMGR_PA_ZIPPATH
$solName      = $env:OOFMGR_PA_SOLNAME
$workflowId   = $env:OOFMGR_PA_WORKFLOWID
$forceFlag    = $env:OOFMGR_PA_FORCE
$connRefsJson = $env:OOFMGR_PA_CONNREFS
$preCreateConn = $env:OOFMGR_PA_PRECREATE_CONN
$cachedInstanceUrl = $env:OOFMGR_PA_CACHED_INSTANCEURL
$cachedEnvId       = $env:OOFMGR_PA_CACHED_ENVID
$cachedEnvDisplay  = $env:OOFMGR_PA_CACHED_ENVDISPLAY
$cachedConnsJson   = $env:OOFMGR_PA_CACHED_CONNS
$cachedFlowMapJson = $env:OOFMGR_PA_CACHED_FLOWMAP
$flowDisplayNamesJson = $env:OOFMGR_PA_FLOWDISPLAYNAMES
$flowDisplayNames = @()
if ($flowDisplayNamesJson) {
    try {
        $parsedFlowNames = ConvertFrom-Json -InputObject $flowDisplayNamesJson
        foreach ($parsedFlowName in @($parsedFlowNames)) {
            foreach ($nameValue in @($parsedFlowName)) {
                if ($nameValue) { $flowDisplayNames += [string]$nameValue }
            }
        }
    } catch {}
}

$debugLog = Join-Path $env:TEMP 'oofmgr-pa-debug.log'
$debugWriter = $null
try {
    $debugFs = [System.IO.FileStream]::new($debugLog, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    $debugWriter = [System.IO.StreamWriter]::new($debugFs, [System.Text.UTF8Encoding]::new($false))
    $debugWriter.AutoFlush = $true
} catch {}
$progressWriter = $null
if ($progressFile) {
    try {
        $progFs = [System.IO.FileStream]::new($progressFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $progressWriter = [System.IO.StreamWriter]::new($progFs, [System.Text.UTF8Encoding]::new($false))
        $progressWriter.AutoFlush = $true
    } catch {}
}
function Trace($m) {
    if ($debugWriter) {
        try { $debugWriter.WriteLine(('{0}  {1}' -f (Get-Date -Format o), $m)) } catch {}
    }
}
function Report($m) {
    if (-not $m) { return }
    Trace (""progress: "" + $m)
    if ($progressWriter) {
        try { $progressWriter.WriteLine([string]$m) } catch {}
    }
}
Trace ""start: pid=$PID op=import upn='$upn' solName='$solName' workflowId='$workflowId' force='$forceFlag' zip='$zipPath'""
Report 'Preparing Power Automate import...'

$Host.UI.RawUI.WindowTitle = 'OofManager - Power Automate import'
Write-Host ''
Write-Host '  OofManager is importing your Cloud Schedule flow into Power Automate.' -ForegroundColor Cyan
Write-Host '  If a sign-in dialog appears, complete it. This window closes automatically.' -ForegroundColor Cyan
Write-Host ''

function Emit($outcome, $message, $envId, $envDisplay, $instanceUrl, $wfid, $conns = $null, $batchItems = @()) {
    $obj = [ordered]@{
        Outcome                = $outcome
        Message                = $message
        EnvironmentId          = $envId
        EnvironmentDisplayName = $envDisplay
        InstanceUrl            = $instanceUrl
        WorkflowId             = $wfid
        Connections            = $conns
        BatchItems             = @($batchItems)
    }
    $json = ($obj | ConvertTo-Json -Compress -Depth 8)
    if ($resultFile) {
        [System.IO.File]::WriteAllText($resultFile, $json, [System.Text.UTF8Encoding]::new($false))
    }
    Trace ""emit: outcome=$outcome message=$message envId=$envId instanceUrl=$instanceUrl wfid=$wfid""
    if ($progressWriter) { try { $progressWriter.Dispose() } catch {} }
    if ($debugWriter) { try { $debugWriter.Dispose() } catch {} }
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
    # Prefer the underlying tools\pac.exe directly over pac.cmd / pac.launcher.exe
    # so each pac invocation skips the extra cmd.exe + launcher process hops.
    $pacPath = $null
    $pacAuthProfilesPath = $null
    if ($env:LOCALAPPDATA) {
        $cliRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\PowerAppsCLI'
        $pacAuthProfilesPath = Join-Path $cliRoot 'authprofiles_v2.json'
        if (Test-Path -LiteralPath $cliRoot) {
            $directPacExe = Get-ChildItem -Path $cliRoot -Filter 'pac.exe' -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | Select-Object -First 1
            if ($directPacExe -and $directPacExe.FullName) { $pacPath = [string]$directPacExe.FullName }
            if (-not $pacPath) {
                $msiPacLauncher = Join-Path $cliRoot 'pac.launcher.exe'
                if (Test-Path -LiteralPath $msiPacLauncher) { $pacPath = $msiPacLauncher }
            }
            if (-not $pacPath) {
                $msiPacCmd = Join-Path $cliRoot 'pac.cmd'
                if (Test-Path -LiteralPath $msiPacCmd) { $pacPath = $msiPacCmd }
            }
        }
    }
    if (-not $pacPath -and $env:USERPROFILE) {
        $dotnetToolPac = Join-Path $env:USERPROFILE '.dotnet\tools\pac.exe'
        if (Test-Path -LiteralPath $dotnetToolPac) { $pacPath = $dotnetToolPac }
    }
    if (-not $pacPath) {
        $pacCmd = Get-Command pac -ErrorAction SilentlyContinue
        if ($pacCmd -and $pacCmd.Source) { $pacPath = [string]$pacCmd.Source }
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
    $zipBytes = (Get-Item -LiteralPath $zipPath).Length
    Trace ""zip size: $zipBytes bytes""

    $script:paSessionReady = $false
    $script:flowJwtToken = $null
    function Ensure-PowerAppsSession {
        if ($script:paSessionReady) { return $true }
        try {
            $ProgressPreference = 'SilentlyContinue'
            $WarningPreference  = 'SilentlyContinue'
            if (-not (Get-Command Get-Flow -ErrorAction SilentlyContinue)) {
                if ($modulesDir -and (Test-Path -LiteralPath $modulesDir) -and ($env:PSModulePath -notlike ""*$modulesDir*"")) {
                    $env:PSModulePath = $modulesDir + ';' + $env:PSModulePath
                }
                Trace 'background prewarm: before Import-Module Microsoft.PowerApps.PowerShell'
                Import-Module Microsoft.PowerApps.PowerShell -DisableNameChecking -ErrorAction Stop | Out-Null
                Trace 'background prewarm: after  Import-Module Microsoft.PowerApps.PowerShell'
            }
            if (-not $global:__oofSignedIn -and (-not $global:currentSession -or $global:currentSession.loggedIn -ne $true)) {
                Trace ""background prewarm: before Add-PowerAppsAccount upn='$upn'""
                if ($upn) { Add-PowerAppsAccount -Endpoint prod -Username $upn -ErrorAction Stop | Out-Null }
                else      { Add-PowerAppsAccount -Endpoint prod -ErrorAction Stop | Out-Null }
                $global:__oofSignedIn = $true
                Trace 'background prewarm: after  Add-PowerAppsAccount'
            }
            try {
                $script:flowJwtToken = Get-JwtToken -Audience 'https://service.flow.microsoft.com/'
            } catch {}
            $script:paSessionReady = $true
            return $true
        } catch {
            Trace ""Ensure-PowerAppsSession failed: $($_.Exception.Message)""
            return $false
        }
    }

    # --- Fast-path: Direct parallel PATCH of existing flows (~8-10s vs ~60s pac solution import) ---
    # When all target flows in this package have already been imported into the
    # user's Developer environment and their runtime FlowName GUIDs are cached,
    # we can update their workflow definitions and set state='Started' directly
    # via the Power Automate ProcessSimple API in parallel. If any flow was
    # deleted in the portal, lacks a required connection reference, or returns
    # an error, we fall through transparently to the full pac solution import.
    if ($envId -and $flowDisplayNames.Count -gt 0 -and $cachedFlowMapJson -and ([string]$cachedFlowMapJson).Trim().Length -gt 2) {
        try {
            $flowMap = @{}
            $parsedMap = $cachedFlowMapJson | ConvertFrom-Json
            if ($parsedMap) {
                foreach ($p in $parsedMap.PSObject.Properties) {
                    if ($p.Name -and $p.Value) { $flowMap[[string]$p.Name] = [string]$p.Value }
                }
            }
            $allFlowsMapped = (@($flowDisplayNames | Where-Object { -not $flowMap.ContainsKey($_) }).Count -eq 0)
            if ($allFlowsMapped) {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
                $zipDefinitions = @{}
                $zipConnRefKeys = @{}
                $za = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
                try {
                    $custEntry = $za.Entries | Where-Object { $_.FullName -ieq 'customizations.xml' } | Select-Object -First 1
                    if ($custEntry) {
                        $csr = New-Object System.IO.StreamReader($custEntry.Open())
                        try { $custXml = [xml]$csr.ReadToEnd() } finally { $csr.Dispose() }
                        foreach ($wfNode in @($custXml.ImportExportXml.Workflows.Workflow)) {
                            if (-not $wfNode) { continue }
                            $wfName = [string]$wfNode.Name
                            $relJson = ([string]$wfNode.JsonFileName).TrimStart('/')
                            $wfEntry = $za.Entries | Where-Object { $_.FullName -ieq $relJson } | Select-Object -First 1
                            if ($wfName -and $wfEntry) {
                                $wsr = New-Object System.IO.StreamReader($wfEntry.Open())
                                try { $wfObj = $wsr.ReadToEnd() | ConvertFrom-Json } finally { $wsr.Dispose() }
                                if ($wfObj -and $wfObj.properties -and $wfObj.properties.definition) {
                                    $zipDefinitions[$wfName] = $wfObj.properties.definition
                                    $reqKeys = @()
                                    if ($wfObj.properties.connectionReferences) {
                                        $reqKeys = @($wfObj.properties.connectionReferences.PSObject.Properties | ForEach-Object { [string]$_.Name })
                                    }
                                    $zipConnRefKeys[$wfName] = $reqKeys
                                }
                            }
                        }
                    }
                } finally {
                    $za.Dispose()
                }

                $allDefsFound = (@($flowDisplayNames | Where-Object { -not $zipDefinitions.ContainsKey($_) }).Count -eq 0)
                if ($allDefsFound -and (Ensure-PowerAppsSession) -and $script:flowJwtToken) {
                    Report ""Updating flow definition(s) in '$envDisplay'...""
                    Trace ""fast PATCH start: env=$envId flows=$($flowDisplayNames -join ' | ')""
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    $flowEndpoint = if ($global:currentSession -and $global:currentSession.flowEndpoint) { [string]$global:currentSession.flowEndpoint } else { 'api.flow.microsoft.com' }
                    $authHeader = if ($script:flowJwtToken -like 'Bearer *') { [string]$script:flowJwtToken } else { 'Bearer ' + [string]$script:flowJwtToken }

                    $getTasks = @()
                    foreach ($dn in $flowDisplayNames) {
                        $fid = $flowMap[$dn]
                        $uri = ""https://$flowEndpoint/providers/Microsoft.ProcessSimple/environments/$envId/flows/$fid`?api-version=2016-11-01""
                        $req = [System.Net.HttpWebRequest]::Create($uri)
                        $req.Method = 'GET'; $req.UserAgent = 'PowerShell cmdlets 1.0'; $req.Headers.Add('Authorization', $authHeader)
                        $getTasks += [pscustomobject]@{ DisplayName = $dn; FlowId = $fid; Uri = $uri; Task = $req.GetResponseAsync() }
                    }

                    $existingFlows = @()
                    $canPatchAll = $true
                    foreach ($gt in $getTasks) {
                        $resp = $gt.Task.GetAwaiter().GetResult()
                        $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
                        try { $fObj = $sr.ReadToEnd() | ConvertFrom-Json } finally { $sr.Dispose(); $resp.Dispose() }
                        if (-not $fObj -or -not $fObj.properties -or ([string]$fObj.properties.displayName -ine $gt.DisplayName)) {
                            Trace ""fast PATCH abort: flow $($gt.FlowId) displayName mismatch or empty""
                            $canPatchAll = $false; break
                        }
                        $existingConnRefs = $fObj.properties.connectionReferences
                        $existingKeys = if ($existingConnRefs) { @($existingConnRefs.PSObject.Properties | ForEach-Object { [string]$_.Name }) } else { @() }
                        foreach ($rk in $zipConnRefKeys[$gt.DisplayName]) {
                            if ($existingKeys -notcontains $rk -or (-not $existingConnRefs.$rk.connectionName)) {
                                Trace ""fast PATCH abort: flow '$($gt.DisplayName)' missing bound connectionReference '$rk'""
                                $canPatchAll = $false; break
                            }
                        }
                        if (-not $canPatchAll) { break }
                        $existingFlows += [pscustomobject]@{ DisplayName = $gt.DisplayName; FlowId = $gt.FlowId; Uri = $gt.Uri; FlowObj = $fObj }
                    }

                    if ($canPatchAll -and $existingFlows.Count -eq $flowDisplayNames.Count) {
                        Trace ""fast PATCH: parallel GET succeeded for $($existingFlows.Count) flow(s); sending parallel PATCH""
                        $patchTasks = @()
                        foreach ($ef in $existingFlows) {
                            $dn = $ef.DisplayName
                            $bodyObj = @{
                                properties = @{
                                    displayName          = [string]$ef.FlowObj.properties.displayName
                                    definition           = $zipDefinitions[$dn]
                                    connectionReferences = $ef.FlowObj.properties.connectionReferences
                                    state                = 'Started'
                                }
                            }
                            $jsonBytes = [System.Text.Encoding]::UTF8.GetBytes(($bodyObj | ConvertTo-Json -Depth 50 -Compress))
                            $req = [System.Net.HttpWebRequest]::Create($ef.Uri)
                            $req.Method = 'PATCH'; $req.ContentType = 'application/json; charset=utf-8'; $req.UserAgent = 'PowerShell cmdlets 1.0'
                            $req.Headers.Add('Authorization', $authHeader); $req.ContentLength = $jsonBytes.Length
                            $reqStream = $req.GetRequestStream()
                            try { $reqStream.Write($jsonBytes, 0, $jsonBytes.Length) } finally { $reqStream.Close() }
                            $patchTasks += [pscustomobject]@{ DisplayName = $dn; FlowId = $ef.FlowId; Task = $req.GetResponseAsync() }
                        }

                        $fastBatchItems = @()
                        foreach ($pt in $patchTasks) {
                            $resp = $pt.Task.GetAwaiter().GetResult()
                            $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
                            try { $resObj = $sr.ReadToEnd() | ConvertFrom-Json } finally { $sr.Dispose(); $resp.Dispose() }
                            $newState = if ($resObj -and $resObj.properties) { [string]$resObj.properties.state } else { '' }
                            Trace ""fast PATCH finished: '$($pt.DisplayName)' flow=$($pt.FlowId) state=$newState""
                            if ($newState -ine 'Started') {
                                Trace ""fast PATCH state not Started ($newState); falling back to pac import""
                                $canPatchAll = $false; break
                            }
                            $refObj = [ordered]@{ EnvironmentName = [string]$envId; FlowName = [string]$pt.FlowId; DisplayName = [string]$pt.DisplayName }
                            $fastBatchItems += [pscustomobject]@{ DisplayName = $pt.DisplayName; Outcome = 'Success'; State = 'On'; Message = 'Power Automate flow is on.'; FlowReference = $refObj }
                        }

                        if ($canPatchAll -and $fastBatchItems.Count -eq $flowDisplayNames.Count) {
                            Report 'Import complete.'
                            Emit 'Success' (""Imported solution '"" + $solName + ""' into '"" + $envDisplay + ""'."" ) $envId $envDisplay $instance $workflowId $null $fastBatchItems
                        }
                    }
                }
            }
        } catch {
            Trace ""fast PATCH path skipped ($($_.Exception.Message)); falling back to pac solution import""
        }
    }

    function Invoke-Pac($stepName, [string[]]$arguments, [scriptblock]$whileRunning = $null) {
        Trace (""pac "" + ($arguments -join ' '))
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $pacPath
        $escapedArgs = @($arguments | ForEach-Object {
            $s = [string]$_
            if ($s -match '[\s""]') { '""' + ($s -replace '""', '\""') + '""' } else { $s }
        })
        $psi.Arguments = ($escapedArgs -join ' ')
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $null = $p.Start()
        $errTask = $p.StandardError.ReadToEndAsync()

        if ($whileRunning) {
            try { & $whileRunning | Out-Null } catch { Trace ""whileRunning callback failed: $($_.Exception.Message)"" }
        }

        while (-not $p.StandardOutput.EndOfStream) {
            $line = $p.StandardOutput.ReadLine()
            if ($null -eq $line) { break }
            Trace (""pac> "" + $line)
            if ($line -match 'Connected to\.\.\.') { Report $line }
            elseif ($line -match 'Solution Importing') { Report ""Importing solution '$solName' into '$envDisplay'..."" }
            elseif ($line -match 'Solution Imported successfully') { Report 'Solution imported. Verifying flow status...' }
            elseif ($line -match 'Publishing All Customizations') { Report 'Publishing Power Automate customizations...' }
            elseif ($line -match 'Published All Customizations') { Report 'Published customizations. Finishing up...' }
            Write-Host $line
        }
        try {
            $errText = $errTask.GetAwaiter().GetResult()
            if ($errText) {
                foreach ($errLine in ($errText -split ""`r?`n"")) {
                    if ($errLine) { Trace (""pac! "" + $errLine) }
                }
            }
        } catch {}
        $p.WaitForExit()
        $code = $p.ExitCode
        $p.Dispose()
        Trace ""pac $stepName exit=$code""
        return $code
    }

    $script:pacAuthEnsured = $false
    function Ensure-PacAuth {
        if ($script:pacAuthEnsured) { return }
        $script:pacAuthEnsured = $true
        $authWhoCode = Invoke-Pac 'auth who' @('auth', 'who')
        if ($authWhoCode -ne 0) {
            # Microsoft corporate accounts use pac's WAM/OperatingSystem profile;
            # creating a browser auth profile can be rejected even when WAM works.
            Trace 'no active pac auth profile; attempting pac auth create'
            $authCode = Invoke-Pac 'auth create' @('auth', 'create', '--environment', $instance)
            if ($authCode -ne 0) {
                Emit 'SignInFailed' (""Automatic Power Automate sign-in failed. Open Power Automate once in your browser, then try again."" ) $envId $envDisplay $instance $null
            }
        }
    }

    # Only run upfront pac auth check when no persisted pac auth profile file
    # exists yet; otherwise skip the ~5s 'pac auth who' process and let the
    # actual connection-list / solution-import command verify auth on demand.
    if (-not ($pacAuthProfilesPath -and (Test-Path -LiteralPath $pacAuthProfilesPath))) {
        Ensure-PacAuth
    } else {
        Trace ""found existing pac auth profile file ($pacAuthProfilesPath); skipping upfront 'pac auth who'""
        Report 'Uploading solution to Power Automate...'
    }

    # --- Auto-bind connection references to existing connections ---------
    # On a fresh tenant the solution's connection references import unbound,
    # so Power Automate refuses to turn the flow on (the 'toggle does nothing
    # until I open it in the portal once' symptom). If the user already has a
    # Connected connection for each required connector in this env, we can
    # bind them at import time with a pac deployment settings file -- no
    # manual step. If any connector has no existing connection (true first
    # run), we skip the settings file and fall back to the manual one-time
    # bind the app already guides the user through.
    $settingsFile = $null
    $connByConnector = @{}
    $usedCachedConns = $false
    if ($connRefsJson -and ([string]$connRefsJson).Trim().Length -gt 2) {
        try {
            $connRefs = $connRefsJson | ConvertFrom-Json
            if ($connRefs -and -not ($connRefs -is [System.Array])) { $connRefs = @($connRefs) }
            Trace ""connRefs requested: $($connRefs.Count)""

            # Reads existing connections in the target env. Prefer the in-process
            # Get-PowerAppConnection cmdlet (~1.2-2s via the already-loaded module)
            # over spawning 'pac connection list' (~7.4s), falling back to pac only
            # if the module call fails. Always query live connections when we reach
            # the full-import path (i.e. fast PATCH was skipped because the flows
            # don't exist yet), so deleted connections are detected and recreated.
            function Get-ConnInventory {
                $connected = @{}
                $present = @{}
                $ok = $false
                if ($envId -and (Ensure-PowerAppsSession)) {
                    try {
                        Trace ""before Get-PowerAppConnection env=$envId""
                        $paConns = @(Get-PowerAppConnection -EnvironmentName $envId -ErrorAction Stop)
                        Trace ""after  Get-PowerAppConnection: count=$($paConns.Count)""
                        foreach ($c in $paConns) {
                            if (-not $c) { continue }
                            $connId = [string]$c.ConnectionName
                            $connector = [string]$c.ConnectorName
                            if (-not $connId -or -not $connector) { continue }
                            $st = $null
                            try { $st = [string]$c.Statuses[0].status } catch {}
                            if (-not $st) { try { $st = [string]$c.Internal.properties.statuses[0].status } catch {} }
                            if (-not $present.ContainsKey($connector)) { $present[$connector] = $connId }
                            if ($st -ieq 'Connected' -and -not $connected.ContainsKey($connector)) {
                                $connected[$connector] = $connId
                            }
                        }
                        $ok = $true
                        return @{ Connected = $connected; Present = $present; Ok = $ok }
                    } catch {
                        Trace ""Get-PowerAppConnection failed ($($_.Exception.Message)); falling back to pac connection list""
                    }
                }

                $lines = @()
                try {
                    $lines = & $pacPath connection list --environment $instance 2>&1 | ForEach-Object { [string]$_ }
                    $ok = ($LASTEXITCODE -eq 0)
                } catch {
                    Trace ""pac connection list failed: $($_.Exception.Message)""
                    $ok = $false
                }
                if (-not $ok -and (-not $script:pacAuthEnsured)) {
                    Ensure-PacAuth
                    try {
                        $lines = & $pacPath connection list --environment $instance 2>&1 | ForEach-Object { [string]$_ }
                        $ok = ($LASTEXITCODE -eq 0)
                    } catch {
                        $ok = $false
                    }
                }
                foreach ($line in $lines) {
                    $apiToken = ($line -split '\s+' | Where-Object { $_ -like '/providers/*/apis/*' } | Select-Object -First 1)
                    if (-not $apiToken) { continue }
                    $tokens = @($line -split '\s+' | Where-Object { $_ -ne '' })
                    if ($tokens.Count -lt 2) { continue }
                    $connId = [string]$tokens[0]
                    $status = [string]$tokens[-1]
                    $connector = ($apiToken -split '/apis/')[-1]
                    if (-not $present.ContainsKey($connector)) { $present[$connector] = $connId }
                    if ($status -ieq 'Connected' -and -not $connected.ContainsKey($connector)) {
                        $connected[$connector] = $connId
                    }
                }
                return @{ Connected = $connected; Present = $present; Ok = $ok }
            }

            # Distinct connectors this solution actually needs.
            $neededConnectors = @($connRefs | ForEach-Object { ($_.connectorId -split '/apis/')[-1] } | Where-Object { $_ } | Select-Object -Unique)

            # Always query the live connection inventory (~1.5s via Get-PowerAppConnection)
            # when we reach the full-import path (meaning the flows do not exist yet or
            # fast PATCH fell through). That way, if the user deleted their connections
            # in Power Automate, $missing is accurately populated and New-OofConnectionInteractive
            # creates fresh connections instead of binding stale cached IDs.
            $inv = Get-ConnInventory
            $connByConnector = $inv.Connected
            foreach ($k in $connByConnector.Keys) { Trace ""connection Connected: connector=$k id=$($connByConnector[$k])"" }
            foreach ($k in $inv.Present.Keys) { Trace ""connection present(any): connector=$k id=$($inv.Present[$k])"" }

            # A connector is 'missing' only when the user has NO connection of
            # that type in ANY status -- not merely because none is currently
            # 'Connected'. An idle office365/flowmanagement connection often
            # lapses to a non-Connected state between vacations; the old
            # Connected-only test treated that as missing and made pre-create
            # mint a DUPLICATE every run. Reuse the existing one instead.
            $missing = @($neededConnectors | Where-Object { -not $inv.Present.ContainsKey($_) })

            # If the list call FAILED we cannot tell what exists, so do NOT
            # pre-create -- minting duplicates is worse than the one-time
            # manual bind the post-import guidance already covers.
            if (-not $inv.Ok) {
                if ($missing.Count -gt 0) { Trace ""connection list unavailable (Ok=false); skipping pre-create to avoid duplicates"" }
                $missing = @()
            }

            # Derive the per-environment Power Platform connectivity host from
            # the environment GUID: strip dashes (32 chars) then split as
            # {first30}.{last2}.environment.api.powerplatform.com. This is the
            # host that serves the getConsentLink / confirmConsentCode actions
            # (verified against the maker portal's own network traffic).
            function Get-PpHost($envGuid) {
                $bare = ([string]$envGuid -replace '-', '')
                if ($bare.Length -lt 3) { return $null }
                return ($bare.Substring(0, $bare.Length - 2) + '.' + $bare.Substring($bare.Length - 2, 2) + '.environment.api.powerplatform.com')
            }

            # POST a per-env connectivity action (getConsentLink /
            # confirmConsentCode). Defined once at script scope (rather than
            # nested inside New-OofConnectionInteractive) so re-defining it on
            # every connector doesn't pollute the outer scope, and so the
            # HttpWebRequest streams can be reliably disposed.
            function Invoke-PpAction($ctx, $action, $bodyObj) {
                $url = ""https://$($ctx.Host)/connectivity/connectors/$($ctx.Connector)/connections/$($ctx.Conn)/$action`?api-version=1&`$filter=$($ctx.Filter)""
                $body = $bodyObj | ConvertTo-Json -Depth 6 -Compress
                $req = [System.Net.HttpWebRequest]::Create($url)
                $req.Method = 'POST'; $req.ContentType = 'application/json'
                $req.Headers.Add('Authorization', ""Bearer $($ctx.Token)"")
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
                $req.ContentLength = $bytes.Length
                $reqStream = $null; $resp = $null; $reader = $null
                try {
                    $reqStream = $req.GetRequestStream()
                    $reqStream.Write($bytes, 0, $bytes.Length)
                    $reqStream.Close()
                    $resp = $req.GetResponse()
                    $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
                    return $reader.ReadToEnd()
                } finally {
                    if ($reader)    { try { $reader.Dispose() }    catch {} }
                    if ($resp)      { try { $resp.Dispose() }      catch {} }
                    if ($reqStream) { try { $reqStream.Dispose() } catch {} }
                }
            }

            # Create ONE user OAuth connection programmatically and authenticate
            # it via the connector's own OAuth consent page shown in an in-app
            # WebView2 dialog. Steps (all verified against live APIs 2026-07-02):
            #   1. PUT a new (unauthenticated) connection on api.powerapps.com.
            #   2. POST getConsentLink on the per-env connectivity host to get
            #      the sign-in URL.
            #   3. Open that URL in WebView2; the user signs in and clicks
            #      ""Allow access"". The apihub finalizes the grant server-side
            #      at its /confirm/Allow step -- there is NO code for us to
            #      redeem (the old confirmConsentCode step is a legacy no-op that
            #      now always 400s).
            #   4. Poll the connection status; when it flips to Connected we
            #      close the window. Connected == success.
            # Returns the connection name after it has reached Connected. Any
            # failure returns $null so the caller can fall back to the maker-
            # portal New-connection page.
            function New-OofConnectionInteractive($connector, $envGuid) {
                # Track the PUT-created connection so any failure path (early
                # return, exception, cancel) can best-effort delete it instead
                # of leaving an orphaned ''OofManager (auto-created)'' unauth
                # connection in the user environment.
                $createdConn = $null
                $success = $false
                try {
                    if (-not (Get-Command Get-JwtToken -ErrorAction SilentlyContinue)) {
                        if ($modulesDir -and (Test-Path -LiteralPath $modulesDir)) { $env:PSModulePath = $modulesDir + ';' + $env:PSModulePath }
                        Import-Module Microsoft.PowerApps.PowerShell -DisableNameChecking -ErrorAction Stop | Out-Null
                    }
                    if (-not $global:__oofSignedIn -and $global:currentSession.loggedIn -ne $true) {
                        if ($upn) { Add-PowerAppsAccount -Endpoint prod -Username $upn -ErrorAction Stop | Out-Null }
                        else      { Add-PowerAppsAccount -Endpoint prod -ErrorAction Stop | Out-Null }
                        $global:__oofSignedIn = $true
                    }

                    $connName = [guid]::NewGuid().ToString('N')

                    # 1. PUT create (api.powerapps.com via module InvokeApi).
                    $createRoute = ""https://{powerAppsEndpoint}/providers/Microsoft.PowerApps/apis/$connector/connections/$connName`?api-version=2016-11-01""
                    $autoName = switch ($connector) {
                        'shared_office365'      { 'OofManager Outlook (auto)' }
                        'shared_flowmanagement' { 'OofManager Flow Management (auto)' }
                        default                 { ""OofManager $connector (auto)"" }
                    }
                    $createBody = @{ properties = @{ environment = @{ id = ""/providers/Microsoft.PowerApps/environments/$envGuid""; name = $envGuid }; displayName = $autoName } }
                    InvokeApi -Method PUT -Route $createRoute -Body $createBody -ApiVersion '2016-11-01' | Out-Null
                    $createdConn = $connName
                    Trace ""created connection $connName for $connector""

                    # Token + host for the per-env connectivity actions.
                    # Switching audience via Add-PowerAppsAccount would wipe the
                    # in-memory session ($global:currentSession = $null) and pop
                    # a SECOND sign-in dialog. Instead reuse the MSAL client we
                    # already signed in above and acquire the api.powerplatform
                    # token SILENTLY (the refresh token is in memory + the
                    # persisted MSAL cache). Only fall back to the interactive
                    # path if silent acquisition genuinely fails, so the normal
                    # case shows just the single connector-consent popup.
                    $ppToken = $null
                    try {
                        if ($global:currentSession -and $global:currentSession.msalClientApp -and $global:currentSession.msalAccount) {
                            [string[]]$ppScopes = 'https://api.powerplatform.com//.default'
                            # Await-Task is private to the module; await the
                            # Task ourselves so we don't depend on it.
                            $ppTask = $global:currentSession.msalClientApp.AcquireTokenSilent($ppScopes, $global:currentSession.msalAccount).ExecuteAsync()
                            # Bound the wait so a stalled network can't block
                            # the whole import; ~10s then treat as a silent
                            # failure and drop to the interactive fallback.
                            $spins = 0
                            while (-not $ppTask.AsyncWaitHandle.WaitOne(200)) { if (++$spins -gt 50) { throw 'silent token acquisition timed out' } }
                            $ar = $ppTask.GetAwaiter().GetResult()
                            $ppToken = $ar.AccessToken
                            Trace 'acquired api.powerplatform.com token silently'
                        }
                    } catch {
                        Trace ""silent api.powerplatform.com token failed: $($_.Exception.Message)""
                    }
                    if (-not $ppToken) {
                        Add-PowerAppsAccount -Audience 'https://api.powerplatform.com/' -ErrorAction Stop | Out-Null
                        $ppToken = Get-JwtToken 'https://api.powerplatform.com/'
                    }
                    $ppHost = Get-PpHost $envGuid
                    if (-not $ppHost -or -not $ppToken) { Trace 'missing pp host/token'; return $false }
                    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue | Out-Null
                    $filter = [System.Web.HttpUtility]::UrlEncode(""environment eq '$envGuid'"")

                    $pp = @{ Host = $ppHost; Token = $ppToken; Connector = $connector; Conn = $connName; Filter = $filter }

                    # 2. getConsentLink. The redirectUrl is where the apihub
                    # sends the browser at the END of the OAuth chain, carrying
                    # the real consent code. Intermediate hops
                    # (login.microsoftonline.com, {connector}.consent.azure-
                    # apim.net) also carry ?code=, but closing the window on one
                    # of those interrupts the apihub's server-side token
                    # exchange before it stores the durable REFRESH token -> the
                    # connection then expires in ~1-2h and any flow using it
                    # fails a few hours later with 'token exchange: missing
                    # connection ACL' (verified 2026-07-02: office365 conn
                    # created 17:38Z, expirationTime 19:05Z). So we only capture
                    # once the browser lands back on THIS host (see below).
                    $redirectUrl = 'https://global.consent.azure-apim.net/redirect'
                    $gcTxt = Invoke-PpAction $pp 'getConsentLink' @{ redirectUrl = $redirectUrl; principalType = 'NotSpecified' }
                    $consent = $gcTxt | ConvertFrom-Json
                    $consentLink = [string]$consent.consentLink
                    if (-not $consentLink) { Trace 'getConsentLink returned no consentLink'; return $false }

                    # 3. Show the sign-in page in an embedded WebView2 (Edge/
                    # Chromium) and capture the redirect's ?code=. The legacy
                    # WinForms WebBrowser (Trident/IE) CANNOT run Microsoft's
                    # modern connector consent confirmation page — its ""Allow
                    # access"" button silently no-ops (submits allowed_plr=False)
                    # so consent never completes and the connection ends up with
                    # only a ~1-2h access token (no refresh) that breaks
                    # unattended flow runs later. WebView2 renders the modern
                    # page correctly so the user can actually grant access ->
                    # durable (~90-day, refreshable) connection. If WebView2 or
                    # its runtime isn't available we return $false and the caller
                    # opens the portal in the real system browser instead.
                    $global:__oofConsentCode = $null
                    $global:__oofWv2Ok = $false
                    $wv2Loaded = $false
                    $wv2Udf = $null
                    try {
                        $wfDll   = Join-Path $appDir 'Microsoft.Web.WebView2.WinForms.dll'
                        $coreDll = Join-Path $appDir 'Microsoft.Web.WebView2.Core.dll'
                        if (-not $appDir -or -not (Test-Path -LiteralPath $wfDll) -or -not (Test-Path -LiteralPath $coreDll)) {
                            Trace ""WebView2 assemblies not found in app dir '$appDir'""
                        } else {
                            # The native WebView2Loader.dll ships next to the exe;
                            # ensure it's on PATH so the managed loader finds it.
                            if ($env:Path -notlike ""*$appDir*"") { $env:Path = $appDir + ';' + $env:Path }
                            $wv2Udf = Join-Path $env:TEMP 'oofmgr-wv2-shared'
                            $env:WEBVIEW2_USER_DATA_FOLDER = $wv2Udf
                            Add-Type -AssemblyName System.Windows.Forms | Out-Null
                            Add-Type -Path $coreDll -ErrorAction Stop | Out-Null
                            Add-Type -Path $wfDll   -ErrorAction Stop | Out-Null
                            $wv2Loaded = $true
                        }
                    } catch {
                        Trace ""WebView2 load failed: $($_.Exception.Message)""
                    }
                    if (-not $wv2Loaded) {
                        Trace 'WebView2 unavailable; returning false so the caller opens the portal in the real browser'
                        return $false
                    }

                    # Create (or reuse across multiple connectors in the same run)
                    # the WebView2 environment with OS-primary-account SSO enabled.
                    if (-not $global:__oofWv2Env) {
                        try {
                            $wv2Opts = New-Object Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions -ArgumentList $null, $null, $null, $true
                            $wv2EnvTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $wv2Udf, $wv2Opts)
                            $spins = 0
                            while (-not $wv2EnvTask.AsyncWaitHandle.WaitOne(200)) { if (++$spins -gt 150) { throw 'environment create timed out' } }
                            $global:__oofWv2Env = $wv2EnvTask.GetAwaiter().GetResult()
                            Trace 'WebView2 environment created (OS-account SSO enabled)'
                        } catch {
                            Trace ""WebView2 environment create failed: $($_.Exception.Message)""
                            return $false
                        }
                    }

                    $form = New-Object System.Windows.Forms.Form
                    $form.Width = 620; $form.Height = 820
                    $form.StartPosition = 'CenterScreen'
                    $form.Text = ""Sign in and click Allow access to create the connection ($connector)""
                    $web = New-Object Microsoft.Web.WebView2.WinForms.WebView2
                    $web.Dock = 'Fill'
                    $form.Controls.Add($web)
                    $global:__oofForm = $form
                    $global:__oofConsentLink = $consentLink
                    $global:__oofUpn = $upn
                    $global:__oofConnected = $false
                    $global:__oofConnector = $connector
                    $global:__oofConnName = $connName
                    $global:__oofEnvGuid = $envGuid
                    $global:__oofPollBusy = $false
                    $global:__oofConsentAllowed = $false

                    # HOW THE CONNECTION ACTUALLY GETS AUTHENTICATED (verified
                    # 2026-07-02 with an isolated harness): it becomes Connected
                    # purely by the user completing sign-in and clicking ""Allow
                    # access"" in THIS window -- the apihub finalizes the grant
                    # server-side at its /confirm/Allow step. There is no code
                    # for us to redeem: the old confirmConsentCode call is a
                    # legacy no-op that now always returns 400, and the previous
                    # code-capture logic closed the window on the FIRST redirect
                    # hop -- which fires BEFORE the user ever reaches the Allow
                    # page -- leaving the connection unauthenticated and sending
                    # us to the portal fallback every time. So we no longer touch
                    # any consent code: we POLL the connection status and close
                    # the window the instant it flips to Connected. Polling
                    # begins only once we're back on an apihub /redirect or
                    # /confirm stage (i.e. past AAD sign-in) so the status calls
                    # don't stutter the password/MFA UI.
                    $global:__oofPollTimer = New-Object System.Windows.Forms.Timer
                    $global:__oofPollTimer.Interval = 1000
                    $global:__oofPollTimer.add_Tick({
                        if ($global:__oofPollBusy) { return }
                        $global:__oofPollBusy = $true
                        try {
                            $c = Get-PowerAppConnection -ConnectorName $global:__oofConnector -ConnectionName $global:__oofConnName -EnvironmentName $global:__oofEnvGuid -ErrorAction Stop
                            if ($c -is [array]) { $c = $c[0] }
                            $st = $null
                            try { $st = [string]$c.Statuses[0].status } catch {}
                            if (-not $st) { try { $st = [string]$c.Internal.properties.statuses[0].status } catch {} }
                            if ($st -eq 'Connected') {
                                $global:__oofConnected = $true
                                $global:__oofPollTimer.Stop()
                                Trace 'connection reached Connected; closing consent window'
                                $global:__oofForm.Close()
                            }
                        } catch {
                        } finally {
                            $global:__oofPollBusy = $false
                        }
                    })
                    # Absolute safety close so the window can never hang if the
                    # user cancels or never finishes the consent.
                    $global:__oofMaxTimer = New-Object System.Windows.Forms.Timer
                    $global:__oofMaxTimer.Interval = 300000
                    $global:__oofMaxTimer.add_Tick({ $global:__oofMaxTimer.Stop(); Trace 'consent window max time reached; closing'; $global:__oofForm.Close() })

                    $web.add_CoreWebView2InitializationCompleted({
                        param($s, $e)
                        if (-not $e.IsSuccess) {
                            Trace ""WebView2 init failed: $($e.InitializationException)""
                            $global:__oofForm.Close(); return
                        }
                        $global:__oofWv2Ok = $true
                        $s.CoreWebView2.add_NavigationStarting({
                            param($s2, $e2)
                            $u = [string]$e2.Uri
                            try { $nav = [Uri]$u; Trace ""consent nav: $($nav.Host)$($nav.AbsolutePath)"" } catch {}
                            if ($global:__oofUpn -and $u -match '^https://login\.(microsoftonline\.com|windows\.net)/[^/]+/oauth2/(v2\.0/)?authorize\?' -and $u -notmatch '[?&]login_hint=') {
                                $hintedUrl = $u + '&login_hint=' + [Uri]::EscapeDataString([string]$global:__oofUpn)
                                Trace 'injecting login_hint into OAuth authorize navigation to skip account picker'
                                $e2.Cancel = $true
                                $s2.Navigate($hintedUrl)
                                return
                            }
                            if ($u -match 'consent\.azure-apim\.net/confirm/Allow') {
                                $global:__oofConsentAllowed = $true
                                $global:__oofForm.Text = ""Authorization complete — finishing connection ($global:__oofConnector)...""
                            }
                            # Both Office 365 and Flow Management currently send
                            # a successful grant through this misleading terminal
                            # ErrorHandler page before the connection flips to
                            # Connected. Suppress only after /confirm/Allow and
                            # only when the URL carries no explicit error value.
                            if ($global:__oofConsentAllowed -and $u -match 'global\.consent\.azure-apim\.net/ErrorHandler' -and $u -notmatch '[?&]error=') {
                                Trace 'suppressing post-Allow consent ErrorHandler navigation while status poll finishes'
                                $e2.Cancel = $true
                                return
                            }
                            if (($u -match 'consent\.azure-apim\.net/redirect' -or $u -match 'consent\.azure-apim\.net/confirm') -and -not $global:__oofPollTimer.Enabled -and -not $global:__oofConnected) {
                                Trace 'reached apihub consent/redirect stage; starting status poll'
                                $global:__oofPollTimer.Start()
                            }
                            if ($u -match '[?&]error=([^&]+)') { Trace ""consent error in redirect: $($matches[1])""; $global:__oofForm.Close() }
                        })
                        $s.CoreWebView2.Navigate($global:__oofConsentLink)
                    })
                    $form.Add_Shown({
                        try { $global:__oofMaxTimer.Start(); $web.EnsureCoreWebView2Async($global:__oofWv2Env) | Out-Null }
                        catch { Trace ""EnsureCoreWebView2Async threw: $($_.Exception.Message)""; $form.Close() }
                    })
                    $form.ShowDialog() | Out-Null
                    try { $global:__oofPollTimer.Stop(); $global:__oofMaxTimer.Stop() } catch {}
                    try { $web.Dispose() } catch {}
                    $form.Dispose()

                    if (-not $global:__oofWv2Ok) { Trace 'WebView2 did not initialize; returning null for portal fallback'; return $null }

                    # Confirm the final status. The poll normally closes the
                    # window the instant it sees Connected (so the flag is set);
                    # re-check once here to cover the max-timeout close path.
                    if (-not $global:__oofConnected) {
                        try {
                            $c = Get-PowerAppConnection -ConnectorName $connector -ConnectionName $connName -EnvironmentName $envGuid -ErrorAction Stop
                            if ($c -is [array]) { $c = $c[0] }
                            $st = $null
                            try { $st = [string]$c.Statuses[0].status } catch {}
                            if (-not $st) { try { $st = [string]$c.Internal.properties.statuses[0].status } catch {} }
                            if ($st -eq 'Connected') { $global:__oofConnected = $true }
                            Trace ""final connection status check: $st""
                        } catch {
                            Trace ""final status check failed: $($_.Exception.Message)""
                        }
                    }

                    if (-not $global:__oofConnected) {
                        Trace ""connection $connName not Connected (cancelled or consent incomplete); returning null for portal fallback""
                        return $null
                    }

                    $success = $true
                    Trace ""connection $connName is Connected""
                    return $connName
                } catch {
                    Trace ""New-OofConnectionInteractive failed for $connector : $($_.Exception.Message)""
                    return $null
                } finally {
                    if (-not $success -and $createdConn) {
                        try {
                            Remove-PowerAppConnection -ConnectorName $connector -ConnectionName $createdConn -EnvironmentName $envGuid -ErrorAction Stop | Out-Null
                            Trace ""cleaned up unauthenticated connection $createdConn for $connector""
                        } catch {
                            Trace ""cleanup of $createdConn failed: $($_.Exception.Message)""
                        }
                    }
                }
            }

            # Pre-create missing connections (opt-in) so the flow can turn on
            # on the FIRST import instead of importing unbound and making the
            # user bind afterward. We create each missing connection in-app and
            # authenticate it with a single sign-in popup (the only browser
            # step). Connectors the user already has are skipped. If the
            # programmatic create fails for any connector, we fall back to the
            # maker-portal New-connection page for that connector. Best-effort:
            # on timeout we fall back to a plain import + post-import guidance.
            if ($preCreateConn -eq '1' -and $envId -and $missing.Count -gt 0) {
                Report ""Creating the required connection(s): $($missing -join ', '). A sign-in window will appear...""
                $pendingCreate = @()
                foreach ($conn in @($missing)) {
                    $createdConnId = New-OofConnectionInteractive $conn $envId
                    if ($createdConnId) {
                        # The helper already observed Connected. Use the exact ID
                        # for deployment settings instead of launching pac again.
                        $connByConnector[$conn] = [string]$createdConnId
                        $inv.Present[$conn] = [string]$createdConnId
                        Trace ""using newly connected connection directly: connector=$conn id=$createdConnId""
                    } else {
                        $pendingCreate += $conn
                        $url = ""https://make.powerautomate.com/environments/$envId/connections/available?apiName=$conn""
                        Trace ""programmatic create failed for $conn; opening portal page: $url""
                        try { Start-Process $url | Out-Null } catch { Trace ""Start-Process failed: $($_.Exception.Message)"" }
                        Report ""Finish creating the '$conn' connection in your browser...""
                    }
                }

                # Only browser-fallback connections need discovery polling. Merge
                # their inventory without replacing exact IDs returned above.
                $deadline = (Get-Date).AddMinutes(4)
                $interval = 5
                while ($pendingCreate.Count -gt 0 -and (Get-Date) -lt $deadline) {
                    $freshInv = Get-ConnInventory
                    foreach ($k in $freshInv.Connected.Keys) { $connByConnector[$k] = $freshInv.Connected[$k] }
                    foreach ($k in $freshInv.Present.Keys) { $inv.Present[$k] = $freshInv.Present[$k] }
                    $pendingCreate = @($pendingCreate | Where-Object { -not $freshInv.Connected.ContainsKey($_) })
                    if ($pendingCreate.Count -eq 0) { break }
                    Report ""Waiting for connection(s) to be ready: $($pendingCreate -join ', ')...""
                    Start-Sleep -Seconds $interval
                    # Back off (5s -> 10s -> 15s -> 20s cap) so a slow consent
                    # doesn't spawn ~48 pac processes over the 4-minute budget.
                    if ($interval -lt 20) { $interval += 5 }
                }
                if ($pendingCreate.Count -eq 0) {
                    Report 'Connection(s) ready. Binding them automatically...'
                } else {
                    Trace ""timed out waiting for connections: $($pendingCreate -join ', '); importing without them""
                }
            }

            # Resolve each requested reference to an existing connection.
            # Bind whatever we CAN (partial binding): the solution may list
            # connectors the user has no connection for yet (true first run),
            # or a reference that isn't even in this package (e.g. the vacation
            # End flow's flow-management ref is omitted when there's no weekly
            # flow to resume). Binding the resolvable refs is strictly better
            # than binding none -- pac binds the listed refs, ignores any not
            # in the solution, and leaves unresolved ones unbound exactly as a
            # plain import would. The app's post-import status check + portal
            # guidance still covers anything left unbound.
            $bindings = @()
            foreach ($cr in $connRefs) {
                $logical = [string]$cr.logical
                $connectorId = [string]$cr.connectorId
                $connectorName = ($connectorId -split '/apis/')[-1]
                $bindConnId = $null
                if ($connByConnector.ContainsKey($connectorName)) {
                    # Prefer a Connected connection.
                    $bindConnId = $connByConnector[$connectorName]
                } elseif ($inv.Present.ContainsKey($connectorName)) {
                    # Fall back to an existing connection that isn't currently
                    # Connected: binding the ref to it (instead of leaving it
                    # unbound or minting a duplicate) means the user only has to
                    # re-consent that one connection to finish setup.
                    $bindConnId = $inv.Present[$connectorName]
                    Trace ""binding ref=$logical to non-Connected existing connection $bindConnId (connector=$connectorName; may need re-consent)""
                }
                if ($bindConnId) {
                    $bindings += [pscustomobject]@{
                        LogicalName = $logical
                        ConnectionId = $bindConnId
                        ConnectorId = $connectorId
                    }
                } else {
                    Trace ""no existing connection for connector=$connectorName (ref=$logical); leaving it unbound""
                }
            }

            if ($bindings.Count -gt 0) {
                $settings = [ordered]@{
                    EnvironmentVariables = @()
                    ConnectionReferences = @($bindings)
                }
                $settingsFile = Join-Path $env:TEMP ('oofmgr-deploy-settings-' + [Guid]::NewGuid().ToString('N') + '.json')
                [System.IO.File]::WriteAllText($settingsFile, ($settings | ConvertTo-Json -Depth 6), [System.Text.UTF8Encoding]::new($false))
                Trace ""wrote settings file with $($bindings.Count) binding(s): $settingsFile""
                Report 'Found existing connections — binding them automatically...'
            } else {
                Trace 'no connection references could be auto-bound; importing without settings file'
            }
        } catch {
            Trace ""auto-bind setup failed (continuing without settings file): $($_.Exception.Message)""
            $settingsFile = $null
        }
    }

    # Omit both --publish-changes and --activate-plugins:
    #   * --publish-changes triggers an org-wide PublishAllXml (~26s).
    #   * --activate-plugins forces Dataverse to deactivate and serially
    #     re-activate already-running flows inside ImportSolution (~19.7s).
    # Instead, we prewarm Microsoft.PowerApps.PowerShell in parallel while
    # pac.exe runs, then verify/enable any Off flows in parallel right here in
    # the same process (~1s when already On, ~12s parallel when Off).
    function Build-ImportArgs($currentSettingsFile) {
        $argsList = @(
            'solution', 'import',
            '--path', $zipPath,
            '--environment', $instance,
            '--max-async-wait-time', '10'
        )
        if ($currentSettingsFile) {
            $argsList += @('--settings-file', $currentSettingsFile)
        }
        if ($forceFlag -eq '1') {
            $argsList += '--force-overwrite'
        }
        return ,$argsList
    }

    $importArgs = Build-ImportArgs $settingsFile
    $importCode = Invoke-Pac 'solution import' $importArgs { Ensure-PowerAppsSession }

    # Self-heal if we skipped upfront auth or used cached connection IDs that
    # became stale (e.g. user deleted and recreated a connection in the portal).
    if ($importCode -ne 0 -and ((-not $script:pacAuthEnsured) -or $usedCachedConns)) {
        Trace 'initial import failed; verifying pac auth and refreshing connection inventory before retry'
        if ($settingsFile) { try { Remove-Item -LiteralPath $settingsFile -ErrorAction SilentlyContinue } catch {} }
        $settingsFile = $null
        Ensure-PacAuth
        if ($usedCachedConns -and $connRefsJson -and ([string]$connRefsJson).Trim().Length -gt 2) {
            try {
                $freshInv = Get-ConnInventory
                $connByConnector = $freshInv.Connected
                $retryBindings = @()
                foreach ($cr in $connRefs) {
                    $logical = [string]$cr.logical
                    $connectorId = [string]$cr.connectorId
                    $connectorName = ($connectorId -split '/apis/')[-1]
                    $bindConnId = if ($connByConnector.ContainsKey($connectorName)) { $connByConnector[$connectorName] }
                                  elseif ($freshInv.Present.ContainsKey($connectorName)) { $freshInv.Present[$connectorName] }
                                  else { $null }
                    if ($bindConnId) {
                        $retryBindings += [pscustomobject]@{ LogicalName = $logical; ConnectionId = $bindConnId; ConnectorId = $connectorId }
                    }
                }
                if ($retryBindings.Count -gt 0) {
                    $settings = [ordered]@{ EnvironmentVariables = @(); ConnectionReferences = @($retryBindings) }
                    $settingsFile = Join-Path $env:TEMP ('oofmgr-deploy-settings-' + [Guid]::NewGuid().ToString('N') + '.json')
                    [System.IO.File]::WriteAllText($settingsFile, ($settings | ConvertTo-Json -Depth 6), [System.Text.UTF8Encoding]::new($false))
                }
            } catch { $settingsFile = $null }
        }
        $importArgs = Build-ImportArgs $settingsFile
        $importCode = Invoke-Pac 'solution import (retry)' $importArgs { Ensure-PowerAppsSession }
    }

    if ($settingsFile) { try { Remove-Item -LiteralPath $settingsFile -ErrorAction SilentlyContinue } catch {} }
    if ($importCode -ne 0) {
        Emit 'ImportFailed' (""Power Automate rejected the solution import for '$solName' in '$envDisplay'."" ) $envId $envDisplay $instance $null
    }

    # Verify and enable imported flows in the same PowerShell process using the
    # already-prewarmed PowerApps session, avoiding a second child process.
    $verifiedBatchItems = @()
    if ($flowDisplayNames.Count -gt 0 -and $envId -and (Ensure-PowerAppsSession)) {
        try {
            function Read-BulkFlowState($flowObj) {
                $res = [pscustomobject]@{ Known = $false; Enabled = $null }
                if (-not $flowObj) { return $res }
                $ep = $flowObj.PSObject.Properties['Enabled']
                $rs = if ($flowObj.PSObject.Properties['Internal'] -and $flowObj.Internal -and $flowObj.Internal.properties) { [string]$flowObj.Internal.properties.state } else { $null }
                if ($ep -ne $null -and $ep.Value -is [bool]) {
                    $agrees = (-not $rs) -or ($rs -ieq 'Started' -and [bool]$ep.Value) -or ($rs -ieq 'Stopped' -and -not [bool]$ep.Value)
                    if ($agrees) { $res.Known = $true; $res.Enabled = [bool]$ep.Value }
                }
                return $res
            }

            Report 'Checking imported flow status...'
            Trace ""in-process post-import check: env=$envId targets=$($flowDisplayNames -join ' | ')""
            $envFlows = @(Get-Flow -EnvironmentName $envId -ErrorAction SilentlyContinue)
            $missingTargets = @($flowDisplayNames | Where-Object { $n = $_; -not ($envFlows | Where-Object { [string]$_.DisplayName -ieq $n }) })
            if ($missingTargets.Count -gt 0) {
                Report 'Waiting for newly imported flows to register...'
                Start-Sleep -Seconds 5
                $envFlows = @(Get-Flow -EnvironmentName $envId -ErrorAction SilentlyContinue)
            }

            $byName = @{}
            $toStart = @()
            foreach ($expectedName in $flowDisplayNames) {
                $f = $envFlows | Where-Object { [string]$_.DisplayName -ieq $expectedName } | Select-Object -First 1
                if (-not $f) {
                    $byName[$expectedName] = [pscustomobject]@{ DisplayName=$expectedName; Outcome='NoFlowFound'; State='NotFound'; Message='Flow was not found after import.'; FlowReference=$null }
                    continue
                }
                $refObj = [ordered]@{ EnvironmentName = [string]$envId; FlowName = [string]$f.FlowName; DisplayName = [string]$f.DisplayName }
                $st = Read-BulkFlowState $f
                if ($st.Known -and $st.Enabled) {
                    Trace ""post-import flow already On: $expectedName flowName=$($f.FlowName)""
                    $byName[$expectedName] = [pscustomobject]@{ DisplayName=$expectedName; Outcome='Success'; State='On'; Message='Power Automate flow is on.'; FlowReference=$refObj }
                } else {
                    $toStart += $f
                }
            }

            if ($toStart.Count -gt 0) {
                $namesLabel = (@($toStart | ForEach-Object { ""'$($_.DisplayName)'"" }) -join ', ')
                Report ""Turning $namesLabel on...""
                $flowEndpoint = if ($global:currentSession -and $global:currentSession.flowEndpoint) { [string]$global:currentSession.flowEndpoint } else { 'api.flow.microsoft.com' }
                $token = $script:flowJwtToken
                if (-not $token) {
                    try {
                        $token = Get-JwtToken -Audience 'https://service.flow.microsoft.com/'
                    } catch {}
                }
                $startErrors = @{}
                if ($token) {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    $authHeader = if ($token -like 'Bearer *') { [string]$token } else { 'Bearer ' + [string]$token }
                    $envSegment = ResolveEnvironment -OverrideId $envId
                    $tasks = @()
                    foreach ($tf in $toStart) {
                        $dn = [string]$tf.DisplayName
                        try {
                            $uri = ""https://$flowEndpoint/providers/Microsoft.ProcessSimple/environments/$envSegment/flows/$($tf.FlowName)/start`?api-version=2016-11-01""
                            Trace ""in-process parallel POST start for '$dn' flow=$($tf.FlowName)""
                            $req = [System.Net.HttpWebRequest]::Create($uri)
                            $req.Method = 'POST'; $req.ContentLength = 0; $req.UserAgent = 'PowerShell cmdlets 1.0'
                            $req.Headers.Add('Authorization', $authHeader)
                            $tasks += [pscustomobject]@{ DisplayName = $dn; Task = $req.GetResponseAsync() }
                        } catch { $startErrors[$dn] = $_.Exception.Message }
                    }
                    foreach ($t in $tasks) {
                        try {
                            $resp = $t.Task.GetAwaiter().GetResult()
                            if ($resp) { try { $resp.Dispose() } catch {} }
                            Trace ""in-process parallel POST start finished for '$($t.DisplayName)'""
                        } catch {
                            $ex = $_.Exception; if ($ex.InnerException) { $ex = $ex.InnerException }
                            $startErrors[$t.DisplayName] = $ex.Message
                        }
                    }
                } else {
                    foreach ($tf in $toStart) {
                        try { Enable-Flow -EnvironmentName $envId -FlowName $tf.FlowName -ErrorAction Stop | Out-Null }
                        catch { $startErrors[[string]$tf.DisplayName] = $_.Exception.Message }
                    }
                }

                for ($vi = 0; $vi -lt 3; $vi++) {
                    $envFlows = @(Get-Flow -EnvironmentName $envId -ErrorAction SilentlyContinue)
                    $allOn = $true
                    foreach ($tf in $toStart) {
                        $dn = [string]$tf.DisplayName
                        if ($startErrors.ContainsKey($dn)) { continue }
                        $rf = $envFlows | Where-Object { [string]$_.DisplayName -ieq $dn } | Select-Object -First 1
                        $st = Read-BulkFlowState $rf
                        if (-not ($st.Known -and $st.Enabled)) { $allOn = $false }
                    }
                    if ($allOn) { break }
                    if ($vi -lt 2) { Start-Sleep -Seconds 3 }
                }

                foreach ($tf in $toStart) {
                    $dn = [string]$tf.DisplayName
                    $refObj = [ordered]@{ EnvironmentName = [string]$envId; FlowName = [string]$tf.FlowName; DisplayName = $dn }
                    if ($startErrors.ContainsKey($dn)) {
                        $byName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='OtherError'; State='Unknown'; Message=$startErrors[$dn]; FlowReference=$refObj }
                        continue
                    }
                    $rf = $envFlows | Where-Object { [string]$_.DisplayName -ieq $dn } | Select-Object -First 1
                    $st = Read-BulkFlowState $rf
                    if ($st.Known -and $st.Enabled) {
                        $byName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='Success'; State='On'; Message='Power Automate flow was turned on.'; FlowReference=$refObj }
                    } elseif (-not $st.Known) {
                        $byName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='Success'; State='Unknown'; Message='Enable was accepted but the final state could not be read.'; FlowReference=$refObj }
                    } else {
                        $byName[$dn] = [pscustomobject]@{ DisplayName=$dn; Outcome='OtherError'; State='Off'; Message='Enable was accepted but the flow is still off (a connection reference is probably unbound).'; FlowReference=$refObj }
                    }
                }
            }

            foreach ($expectedName in $flowDisplayNames) {
                if ($byName.ContainsKey($expectedName)) { $verifiedBatchItems += $byName[$expectedName] }
            }
        } catch {
            Trace ""in-process post-import check failed (caller will fallback): $($_.Exception.Message)""
            $verifiedBatchItems = @()
        }
    }

    Report 'Import complete.'
    Emit 'Success' (""Imported solution '"" + $solName + ""' into '"" + $envDisplay + ""'."" ) $envId $envDisplay $instance $workflowId $connByConnector $verifiedBatchItems

} catch {
    Trace ""OUTER CATCH: $($_.Exception.Message)""
    Emit 'OtherError' $_.Exception.Message $null $null $null $null
}
";
}
