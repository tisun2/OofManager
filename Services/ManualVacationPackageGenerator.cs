using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.Json;

namespace OofManager.Wpf.Services;

/// <summary>
/// Generates a Dataverse <em>Solution</em> .zip that, when imported under
/// <em>Power Automate &rarr; Solutions &rarr; Import solution</em>, creates
/// <strong>two</strong> one-shot cloud flows that orchestrate a manual
/// vacation entirely from Microsoft 365:
/// <list type="number">
///   <item><description><b>Vacation Start</b> — fires once at the user's
///   vacation start time. Sets Outlook automatic replies (scheduled, with
///   both start/end), then turns OFF the user's existing
///   <em>OofManager Cloud Schedule</em> flow so the weekly schedule doesn't
///   keep flipping OOF on/off during the vacation.</description></item>
///   <item><description><b>Vacation End</b> — fires once at the user's
///   vacation end time. Turns the Cloud Schedule flow back ON. Outlook
///   automatically clears the AutoReply at the scheduledEndDateTime set by
///   the Start flow, so we don't need a second AutoReply action here.</description></item>
/// </list>
/// Companion to <see cref="CloudSchedulePackageGenerator"/>; identifiers
/// are derived from the signed-in user's alias the same way so re-importing
/// after the user changes their vacation window upgrades the same solution
/// in place rather than creating duplicates.
/// </summary>
public static class ManualVacationPackageGenerator
{
    private const string PublisherUniqueName = "OofManagerPublisher";
    private const string PublisherDisplayName = "OofManager";
    private const string PublisherPrefix = "ofm";
    private const int PublisherCustomizationOption = 10000;

    // Namespaces for deterministic v5 GUIDs. Disjoint from the
    // CloudSchedulePackageGenerator namespaces so the same alias produces
    // a different workflow id here vs. the Cloud Schedule flow — otherwise
    // a re-import of either solution could clobber the other's workflow row.
    private static readonly Guid VacationStartWorkflowNamespace = new("a31f7e64-2c8d-4f9e-b1c2-3d4e5f6a7b81");
    private static readonly Guid VacationEndWorkflowNamespace   = new("b42e8f75-3d9e-4a0f-c2d3-4e5f6a7b8c92");
    private static readonly Guid OutlookConnRefNamespace        = new("c530a986-4eaf-4b1f-d3e4-5f6a7b8c9da3");
    private static readonly Guid FlowMgmtConnRefNamespace       = new("d641ba97-5fb0-4c2e-e4f5-6a7b8c9daeb4");

    private const string OutlookConnReferenceDisplayName  = "OofManager Outlook";
    private const string FlowMgmtConnReferenceDisplayName = "OofManager Flow Management";
    private const string OutlookConnectorId  = "/providers/Microsoft.PowerApps/apis/shared_office365";
    private const string FlowMgmtConnectorId = "/providers/Microsoft.PowerApps/apis/shared_flowmanagement";
    private const int SolutionMajorVersion = 1;

    /// <summary>
    /// Per-user/per-vacation identifiers stamped into the solution package.
    /// </summary>
    private sealed class VacationIdentity
    {
        public string Alias { get; }
        public string SolutionUniqueName { get; }
        public string SolutionDisplayName { get; }
        public string StartFlowDisplayName { get; }
        public string EndFlowDisplayName { get; }
        public string StartWorkflowFileName { get; }
        public string EndWorkflowFileName { get; }
        public Guid StartWorkflowId { get; }
        public Guid EndWorkflowId { get; }
        public Guid OutlookConnRefId { get; }
        public Guid FlowMgmtConnRefId { get; }
        public string OutlookConnRefLogicalName { get; }
        public string FlowMgmtConnRefLogicalName { get; }

        public VacationIdentity(
            string alias,
            string solutionUniqueName,
            string solutionDisplayName,
            string startFlowDisplayName,
            string endFlowDisplayName,
            string startWorkflowFileName,
            string endWorkflowFileName,
            Guid startWorkflowId,
            Guid endWorkflowId,
            Guid outlookConnRefId,
            Guid flowMgmtConnRefId,
            string outlookConnRefLogicalName,
            string flowMgmtConnRefLogicalName)
        {
            Alias = alias;
            SolutionUniqueName = solutionUniqueName;
            SolutionDisplayName = solutionDisplayName;
            StartFlowDisplayName = startFlowDisplayName;
            EndFlowDisplayName = endFlowDisplayName;
            StartWorkflowFileName = startWorkflowFileName;
            EndWorkflowFileName = endWorkflowFileName;
            StartWorkflowId = startWorkflowId;
            EndWorkflowId = endWorkflowId;
            OutlookConnRefId = outlookConnRefId;
            FlowMgmtConnRefId = flowMgmtConnRefId;
            OutlookConnRefLogicalName = outlookConnRefLogicalName;
            FlowMgmtConnRefLogicalName = flowMgmtConnRefLogicalName;
        }
    }

    /// <summary>
    /// Result of <see cref="GenerateWithIdentity"/> — path of the zip on
    /// disk plus the per-user identifiers baked into it.
    /// </summary>
    public sealed class GenerateResult
    {
        public GenerateResult(
            string path,
            string solutionUniqueName,
            string solutionVersion,
            string startFlowDisplayName,
            string endFlowDisplayName,
            Guid startWorkflowId,
            Guid endWorkflowId)
        {
            Path = path;
            SolutionUniqueName = solutionUniqueName;
            SolutionVersion = solutionVersion;
            StartFlowDisplayName = startFlowDisplayName;
            EndFlowDisplayName = endFlowDisplayName;
            StartWorkflowId = startWorkflowId;
            EndWorkflowId = endWorkflowId;
        }

        public string Path { get; }
        public string SolutionUniqueName { get; }
        public string SolutionVersion { get; }
        public string StartFlowDisplayName { get; }
        public string EndFlowDisplayName { get; }
        public Guid StartWorkflowId { get; }
        public Guid EndWorkflowId { get; }
    }

    /// <summary>
    /// Returns the deterministic per-user identifiers without generating a
    /// package. Lets callers — e.g. "Cancel planned vacation" — look up the
    /// solution unique name by alias alone, without re-deriving it.
    /// </summary>
    public static GenerateResult ComputeIdentity(string userEmail)
    {
        var i = BuildIdentity(userEmail);
        return new GenerateResult(
            path: string.Empty,
            solutionUniqueName: i.SolutionUniqueName,
            solutionVersion: string.Empty,
            startFlowDisplayName: i.StartFlowDisplayName,
            endFlowDisplayName: i.EndFlowDisplayName,
            startWorkflowId: i.StartWorkflowId,
            endWorkflowId: i.EndWorkflowId);
    }

    /// <summary>
    /// Builds the package and writes it to <paramref name="outputPath"/>.
    /// Pass null to drop it under <c>%TEMP%\OofManager-ManualVacation.zip</c>.
    /// </summary>
    /// <param name="userEmail">UPN of the signed-in user; used to derive the
    /// per-user solution unique name and deterministic workflow GUIDs.</param>
    /// <param name="vacationStart">When the AutoReply should turn on and the
    /// Schedule flow should be paused. Local kind acceptable — converted to
    /// UTC for the Recurrence trigger.</param>
    /// <param name="vacationEnd">When the Schedule flow should be re-enabled.
    /// Also used as <c>scheduledEndDateTime</c> on the AutoReply so Outlook
    /// clears the reply automatically.</param>
    /// <param name="internalReply">Internal AutoReply body (plain text;
    /// wrapped in HTML the same way the Cloud Schedule generator does).</param>
    /// <param name="externalReply">External AutoReply body (plain text).</param>
    /// <param name="externalAudienceAll">Send external reply to everyone if
    /// true; only known contacts if false. Matches <c>OofSettings.ExternalAudienceAll</c>.</param>
    /// <param name="scheduleFlowEnvironmentId">Environment id of the existing
    /// OofManager Cloud Schedule flow — used to target the
    /// pause/resume actions. Pulled from the Power Automate import-environment
    /// cache. Pass null/empty to omit the pause/resume actions (Phase-1 debug
    /// drop where only AutoReply is wired).</param>
    /// <param name="scheduleFlowRuntimeFlowName">Runtime <c>FlowName</c> GUID
    /// (NOT the deterministic WorkflowId from the Cloud Schedule package — see
    /// repo memory note on <c>FlowName</c> != WorkflowId after import) of the
    /// Schedule flow. Pulled from the Power Automate flow-reference cache.
    /// Pass null/empty to omit pause/resume actions.</param>
    public static GenerateResult GenerateWithIdentity(
        string userEmail,
        DateTime vacationStart,
        DateTime vacationEnd,
        string internalReply,
        string externalReply,
        bool externalAudienceAll,
        string? scheduleFlowEnvironmentId,
        string? scheduleFlowRuntimeFlowName,
        bool generateManaged = true,
        string? outputPath = null)
    {
        if (vacationEnd <= vacationStart)
            throw new ArgumentException("vacationEnd must be after vacationStart.", nameof(vacationEnd));

        outputPath ??= Path.Combine(Path.GetTempPath(), "OofManager-ManualVacation.zip");

        var identity = BuildIdentity(userEmail);
        var generatedAt = DateTime.Now;
        var solutionVersion = BuildSolutionVersion(generatedAt);

        var tzId = TimeZoneInfo.Local.Id;

        // Recurrence triggers want startTime in ISO 8601. We convert to UTC
        // and serialise with a trailing 'Z' so the trigger fires at the same
        // instant the user picked locally, regardless of where the Power
        // Automate runner happens to evaluate the schedule.
        var startUtc = ToUtcIso8601(vacationStart);
        var endUtc   = ToUtcIso8601(vacationEnd);

        // The AutoReply action wants scheduledStart/End as naive strings
        // accompanied by an explicit timeZone field; pass the user's local
        // wall-clock so Outlook displays the reply window exactly as the
        // user typed it.
        var autoReplyStart = vacationStart.ToString("yyyy-MM-ddTHH:mm:ss");
        var autoReplyEnd   = vacationEnd.ToString("yyyy-MM-ddTHH:mm:ss");

        var hasScheduleFlowTarget =
            !string.IsNullOrWhiteSpace(scheduleFlowEnvironmentId) &&
            !string.IsNullOrWhiteSpace(scheduleFlowRuntimeFlowName);

        var startFlowJson = BuildVacationStartFlowJson(
            identity: identity,
            tzId: tzId,
            triggerStartUtc: startUtc,
            autoReplyStart: autoReplyStart,
            autoReplyEnd: autoReplyEnd,
            audience: externalAudienceAll ? "all" : "contactsOnly",
            internalReply: PlainTextToHtml(internalReply),
            externalReply: PlainTextToHtml(externalReply),
            scheduleFlowEnvironmentId: scheduleFlowEnvironmentId,
            scheduleFlowRuntimeFlowName: scheduleFlowRuntimeFlowName);

        var endFlowJson = BuildVacationEndFlowJson(
            identity: identity,
            triggerStartUtc: endUtc,
            scheduleFlowEnvironmentId: scheduleFlowEnvironmentId,
            scheduleFlowRuntimeFlowName: scheduleFlowRuntimeFlowName);

        var solutionXml = BuildSolutionXml(generateManaged, identity, solutionVersion);
        var customizationsXml = BuildCustomizationsXml(identity, hasScheduleFlowTarget);
        var contentTypesXml = BuildContentTypesXml();

        if (File.Exists(outputPath))
        {
            try { File.Delete(outputPath); } catch { /* best-effort */ }
        }
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        using (var fs = File.Create(outputPath))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            WriteEntry(zip, "[Content_Types].xml", contentTypesXml);
            WriteEntry(zip, "solution.xml", solutionXml);
            WriteEntry(zip, "customizations.xml", customizationsXml);
            WriteEntry(zip, $"Workflows/{identity.StartWorkflowFileName}", startFlowJson);
            WriteEntry(zip, $"Workflows/{identity.EndWorkflowFileName}", endFlowJson);
            WriteEntry(zip, "OofManager-README.txt", BuildReadme(userEmail, identity, vacationStart, vacationEnd, hasScheduleFlowTarget));
        }

        return new GenerateResult(
            path: outputPath,
            solutionUniqueName: identity.SolutionUniqueName,
            solutionVersion: solutionVersion,
            startFlowDisplayName: identity.StartFlowDisplayName,
            endFlowDisplayName: identity.EndFlowDisplayName,
            startWorkflowId: identity.StartWorkflowId,
            endWorkflowId: identity.EndWorkflowId);
    }

    private static string BuildSolutionVersion(DateTime generatedAt)
    {
        var minor = generatedAt.Year % 100;
        var build = generatedAt.Month * 100 + generatedAt.Day;
        var revision = generatedAt.Hour * 100 + generatedAt.Minute;
        return $"{SolutionMajorVersion}.{minor}.{build}.{revision}";
    }

    private static string BuildWorkflowFileName(string flowDisplayName, Guid workflowId) =>
        $"{flowDisplayName.Replace(' ', '_')}-{workflowId.ToString("D").ToUpperInvariant()}.json";

    private static string ToUtcIso8601(DateTime local)
    {
        var utc = local.Kind == DateTimeKind.Utc ? local : local.ToUniversalTime();
        return utc.ToString("yyyy-MM-ddTHH:mm:ssZ");
    }

    private static void WriteEntry(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    /// <summary>
    /// Builds the vacation-start workflow JSON. One-shot Recurrence trigger at
    /// the user's vacation start. When a Schedule-flow target is supplied the
    /// action chain is ordered to defeat the SetAutoReply race against the
    /// Schedule flow (which may fire at the exact same instant — e.g. 18:00
    /// Monday — and otherwise overwrite the vacation reply with its daily OOF
    /// text on the way past), without leaving a window where the vacation
    /// reply isn't live yet (matters when the vacation starts mid-day while
    /// Weekly OOF is OFF — without this, clients emailing during the wait
    /// would get no auto-reply at all):
    /// <list type="number">
    ///   <item><description><b>StopFlow</b> Schedule — prevents future
    ///   Schedule-flow triggers during the vacation.</description></item>
    ///   <item><description><b>SetAutoReply (first write)</b> — vacation
    ///   reply goes live immediately so anyone emailing during the next
    ///   60 seconds gets the vacation text.</description></item>
    ///   <item><description><b>Delay 60 seconds</b> — lets any Schedule-flow
    ///   run that was already in-flight at T0 finish its own SetAutoReply
    ///   (which would otherwise overwrite our first write). Schedule runs
    ///   normally finish in 5–30 s; 60 s covers the vast majority and any
    ///   rare slow run only briefly clobbers our first write before the
    ///   second write below restores it.</description></item>
    ///   <item><description><b>SetAutoReply (second write)</b> — last write
    ///   wins; if the in-flight Schedule run clobbered our first write, this
    ///   restores the vacation reply.</description></item>
    /// </list>
    /// When no Schedule-flow target is supplied (no cached runtime FlowName),
    /// only a single SetAutoReply runs — there's nothing to race against.
    /// </summary>
    private static string BuildVacationStartFlowJson(
        VacationIdentity identity,
        string tzId,
        string triggerStartUtc,
        string autoReplyStart,
        string autoReplyEnd,
        string audience,
        string internalReply,
        string externalReply,
        string? scheduleFlowEnvironmentId,
        string? scheduleFlowRuntimeFlowName)
    {
        var hasScheduleTarget =
            !string.IsNullOrWhiteSpace(scheduleFlowEnvironmentId) &&
            !string.IsNullOrWhiteSpace(scheduleFlowRuntimeFlowName);

        var connectionReferences = new Dictionary<string, object?>
        {
            ["shared_office365"] = BuildConnectionReferenceMap("shared_office365", identity.OutlookConnRefLogicalName),
        };

        var actions = new Dictionary<string, object?>();

        if (hasScheduleTarget)
        {
            connectionReferences["shared_flowmanagement"] =
                BuildConnectionReferenceMap("shared_flowmanagement", identity.FlowMgmtConnRefLogicalName);

            // 1) Pause the Schedule flow FIRST so no future recurrence fires.
            actions["Pause_OofManager_Cloud_Schedule_flow"] = BuildTurnOffFlowAction(
                scheduleFlowEnvironmentId!, scheduleFlowRuntimeFlowName!,
                runAfterStep: null);

            // 2) Vacation SetAutoReply (FIRST write) — gives anyone emailing
            //    during the upcoming 60-second wait the vacation reply right
            //    away, instead of falling back to whatever Outlook had
            //    (possibly OFF, if the vacation starts mid-day).
            actions["Set_up_automatic_replies_first"] = BuildSetAutoReplyAction(
                tzId: tzId,
                autoReplyStart: autoReplyStart,
                autoReplyEnd: autoReplyEnd,
                audience: audience,
                internalReply: internalReply,
                externalReply: externalReply,
                runAfterStep: "Pause_OofManager_Cloud_Schedule_flow");

            // 3) Wait 60 seconds for any in-flight Schedule-flow run that was
            //    triggered at the same instant T0 to finish its own
            //    SetAutoReply. Built-in 'Wait' action — no connector reference
            //    needed. Schedule runs normally finish in 5–30 s.
            actions["Wait_for_in_flight_Schedule_run"] = new Dictionary<string, object?>
            {
                ["runAfter"] = new Dictionary<string, object?>
                {
                    ["Set_up_automatic_replies_first"] = new[] { "Succeeded", "Failed", "Skipped", "TimedOut" },
                },
                ["type"] = "Wait",
                ["inputs"] = new Dictionary<string, object?>
                {
                    ["interval"] = new Dictionary<string, object?>
                    {
                        ["count"] = 60,
                        ["unit"] = "Second",
                    },
                },
            };

            // 4) Vacation SetAutoReply (SECOND write) — last write wins. If
            //    the in-flight Schedule run clobbered our first write during
            //    the wait, this restores the vacation reply as the final
            //    state Outlook sees.
            actions["Set_up_automatic_replies"] = BuildSetAutoReplyAction(
                tzId: tzId,
                autoReplyStart: autoReplyStart,
                autoReplyEnd: autoReplyEnd,
                audience: audience,
                internalReply: internalReply,
                externalReply: externalReply,
                runAfterStep: "Wait_for_in_flight_Schedule_run");
        }
        else
        {
            // No Schedule target → no race possible → single action.
            actions["Set_up_automatic_replies"] = BuildSetAutoReplyAction(
                tzId: tzId,
                autoReplyStart: autoReplyStart,
                autoReplyEnd: autoReplyEnd,
                audience: audience,
                internalReply: internalReply,
                externalReply: externalReply,
                runAfterStep: null);
        }

        return SerializeWorkflowEnvelope(
            displayName: identity.StartFlowDisplayName,
            connectionReferences: connectionReferences,
            trigger: BuildOneShotRecurrence(triggerStartUtc),
            actions: actions);
    }

    /// <summary>
    /// Builds the vacation-end workflow JSON. One-shot Recurrence trigger at
    /// vacation end; single action turning the Cloud Schedule flow back on.
    /// Omitted entirely (zero actions) when no Schedule-flow target is
    /// supplied, in which case the flow becomes a no-op — Outlook will still
    /// clear the AutoReply on its own via the scheduledEndDateTime set by
    /// the start flow.
    /// </summary>
    private static string BuildVacationEndFlowJson(
        VacationIdentity identity,
        string triggerStartUtc,
        string? scheduleFlowEnvironmentId,
        string? scheduleFlowRuntimeFlowName)
    {
        var connectionReferences = new Dictionary<string, object?>();
        var actions = new Dictionary<string, object?>();

        if (!string.IsNullOrWhiteSpace(scheduleFlowEnvironmentId) &&
            !string.IsNullOrWhiteSpace(scheduleFlowRuntimeFlowName))
        {
            connectionReferences["shared_flowmanagement"] =
                BuildConnectionReferenceMap("shared_flowmanagement", identity.FlowMgmtConnRefLogicalName);
            actions["Resume_OofManager_Weekly_Schedule_flow"] = BuildTurnOnFlowAction(
                scheduleFlowEnvironmentId!, scheduleFlowRuntimeFlowName!,
                runAfterStep: null);
        }

        return SerializeWorkflowEnvelope(
            displayName: identity.EndFlowDisplayName,
            connectionReferences: connectionReferences,
            trigger: BuildOneShotRecurrence(triggerStartUtc),
            actions: actions);
    }

    private static Dictionary<string, object?> BuildConnectionReferenceMap(string apiName, string logicalName) =>
        new()
        {
            ["runtimeSource"] = "embedded",
            ["connection"] = new Dictionary<string, object?>
            {
                ["connectionReferenceLogicalName"] = logicalName,
            },
            ["api"] = new Dictionary<string, object?>
            {
                ["name"] = apiName,
            },
        };

    private static Dictionary<string, object?> BuildOneShotRecurrence(string startTimeUtcIso) =>
        new()
        {
            // One-shot pattern: a Month×1 recurrence starting at the chosen
            // instant with count=1. We previously used Year×1, but the new
            // Power Automate designer's Frequency dropdown only lists
            // Second/Minute/Hour/Day/Week/Month — "Year" is accepted by the
            // engine but renders blank ("Select frequency.") in the UI, which
            // looks broken and risks the required field being cleared if a
            // user opens+saves the flow. Month×1 (≈30 days) displays correctly,
            // stays well under the 500-day frequency×interval cap, and count=1
            // still pins it to exactly one run at startTime (the frequency
            // value is irrelevant to that single fire).
            ["recurrence"] = new Dictionary<string, object?>
            {
                ["frequency"] = "Month",
                ["interval"] = 1,
                ["startTime"] = startTimeUtcIso,
                ["timeZone"] = "UTC",
                ["count"] = 1,
            },
            ["type"] = "Recurrence",
        };

    private static Dictionary<string, object?> BuildSetAutoReplyAction(
        string tzId,
        string autoReplyStart,
        string autoReplyEnd,
        string audience,
        string internalReply,
        string externalReply,
        string? runAfterStep)
    {
        var runAfter = new Dictionary<string, object?>();
        if (!string.IsNullOrWhiteSpace(runAfterStep))
        {
            // Run AutoReply even if the preceding Wait/Pause reported a
            // non-Succeeded state — we still want vacation reply set.
            runAfter[runAfterStep!] = new[] { "Succeeded", "Failed", "Skipped", "TimedOut" };
        }
        return new Dictionary<string, object?>
        {
            ["runAfter"] = runAfter,
            ["metadata"] = new Dictionary<string, object?>
            {
                ["operationMetadataId"] = Guid.NewGuid().ToString("D"),
            },
            ["type"] = "OpenApiConnection",
            ["inputs"] = new Dictionary<string, object?>
            {
                ["host"] = new Dictionary<string, object?>
                {
                    ["connectionName"] = "shared_office365",
                    ["operationId"] = "SetAutomaticRepliesSetting_V2",
                    ["apiId"] = "/providers/Microsoft.PowerApps/apis/shared_office365",
                },
                ["parameters"] = new Dictionary<string, object?>
                {
                    ["body/automaticRepliesSetting/status"] = "scheduled",
                    ["body/automaticRepliesSetting/externalAudience"] = audience,
                    ["body/automaticRepliesSetting/scheduledStartDateTime/dateTime"] = autoReplyStart,
                    ["body/automaticRepliesSetting/scheduledStartDateTime/timeZone"] = tzId,
                    ["body/automaticRepliesSetting/scheduledEndDateTime/dateTime"] = autoReplyEnd,
                    ["body/automaticRepliesSetting/scheduledEndDateTime/timeZone"] = tzId,
                    ["body/automaticRepliesSetting/internalReplyMessage"] = internalReply ?? string.Empty,
                    ["body/automaticRepliesSetting/externalReplyMessage"] = externalReply ?? string.Empty,
                },
                ["authentication"] = "@parameters('$authentication')",
            },
        };
    }

    private static Dictionary<string, object?> BuildTurnOffFlowAction(string envId, string flowName, string? runAfterStep)
        => BuildFlowMgmtAction("StopFlow", envId, flowName, runAfterStep);

    private static Dictionary<string, object?> BuildTurnOnFlowAction(string envId, string flowName, string? runAfterStep)
        => BuildFlowMgmtAction("StartFlow", envId, flowName, runAfterStep);

    /// <summary>
    /// Builds a Power Automate Management connector action. operationId
    /// <c>StopFlow</c> / <c>StartFlow</c> are the operation ids exposed by
    /// the <c>shared_flowmanagement</c> connector for "Turn off flow" /
    /// "Turn on flow" respectively. The connector's parameters surface
    /// <em>environmentName</em> + <em>flowName</em> (the runtime workflow
    /// GUID).
    /// </summary>
    private static Dictionary<string, object?> BuildFlowMgmtAction(
        string operationId,
        string envId,
        string flowName,
        string? runAfterStep)
    {
        var runAfter = new Dictionary<string, object?>();
        if (!string.IsNullOrWhiteSpace(runAfterStep))
        {
            runAfter[runAfterStep!] = new[] { "Succeeded" };
        }

        return new Dictionary<string, object?>
        {
            ["runAfter"] = runAfter,
            ["metadata"] = new Dictionary<string, object?>
            {
                ["operationMetadataId"] = Guid.NewGuid().ToString("D"),
            },
            ["type"] = "OpenApiConnection",
            ["inputs"] = new Dictionary<string, object?>
            {
                ["host"] = new Dictionary<string, object?>
                {
                    ["connectionName"] = "shared_flowmanagement",
                    ["operationId"] = operationId,
                    ["apiId"] = "/providers/Microsoft.PowerApps/apis/shared_flowmanagement",
                },
                ["parameters"] = new Dictionary<string, object?>
                {
                    ["environmentName"] = envId,
                    ["flowName"] = flowName,
                },
                ["authentication"] = "@parameters('$authentication')",
            },
        };
    }

    private static string SerializeWorkflowEnvelope(
        string displayName,
        Dictionary<string, object?> connectionReferences,
        Dictionary<string, object?> trigger,
        Dictionary<string, object?> actions)
    {
        var wrapper = new Dictionary<string, object?>
        {
            ["properties"] = new Dictionary<string, object?>
            {
                ["connectionReferences"] = connectionReferences,
                ["definition"] = new Dictionary<string, object?>
                {
                    ["$schema"] = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
                    ["contentVersion"] = "1.0.0.0",
                    ["parameters"] = new Dictionary<string, object?>
                    {
                        ["$connections"] = new Dictionary<string, object?>
                        {
                            ["defaultValue"] = new Dictionary<string, object?>(),
                            ["type"] = "Object",
                        },
                        ["$authentication"] = new Dictionary<string, object?>
                        {
                            ["defaultValue"] = new Dictionary<string, object?>(),
                            ["type"] = "SecureObject",
                        },
                    },
                    ["triggers"] = new Dictionary<string, object?>
                    {
                        ["Recurrence"] = trigger,
                    },
                    ["actions"] = actions,
                },
                ["parameters"] = new Dictionary<string, object?>(),
                ["displayName"] = displayName,
            },
            ["schemaVersion"] = "1.0.0.0",
        };

        return JsonSerializer.Serialize(wrapper, JsonOpts);
    }

    private static string BuildSolutionXml(bool managed, VacationIdentity identity, string solutionVersion)
    {
        // Two RootComponent entries (one per workflow), both component type 29.
        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<ImportExportXml version=""9.2.0.1234"" SolutionPackageVersion=""9.2"" languagecode=""1033"" generatedBy=""OofManager"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <SolutionManifest>
    <UniqueName>{identity.SolutionUniqueName}</UniqueName>
    <LocalizedNames>
      <LocalizedName description=""{XmlEscape(identity.SolutionDisplayName)}"" languagecode=""1033"" />
    </LocalizedNames>
    <Descriptions>
      <Description description=""Manual-vacation cloud orchestrator generated by OofManager. Sets Outlook AutoReply at the chosen start and re-enables the OofManager weekly schedule at the chosen end."" languagecode=""1033"" />
    </Descriptions>
    <Version>{solutionVersion}</Version>
    <Managed>{(managed ? 1 : 0)}</Managed>
    <Publisher>
      <UniqueName>{PublisherUniqueName}</UniqueName>
      <LocalizedNames>
        <LocalizedName description=""{XmlEscape(PublisherDisplayName)}"" languagecode=""1033"" />
      </LocalizedNames>
      <Descriptions>
        <Description description=""Default publisher for OofManager-generated solutions."" languagecode=""1033"" />
      </Descriptions>
      <EMailAddress xsi:nil=""true""></EMailAddress>
      <SupportingWebsiteUrl xsi:nil=""true""></SupportingWebsiteUrl>
      <CustomizationPrefix>{PublisherPrefix}</CustomizationPrefix>
      <CustomizationOptionValuePrefix>{PublisherCustomizationOption}</CustomizationOptionValuePrefix>
      <Addresses>
        <Address>
          <AddressNumber>1</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          <ShippingMethodCode>1</ShippingMethodCode>
        </Address>
        <Address>
          <AddressNumber>2</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          <ShippingMethodCode>1</ShippingMethodCode>
        </Address>
      </Addresses>
    </Publisher>
    <RootComponents>
      <RootComponent type=""29"" id=""{{{identity.StartWorkflowId.ToString("D").ToLowerInvariant()}}}"" behavior=""0"" />
      <RootComponent type=""29"" id=""{{{identity.EndWorkflowId.ToString("D").ToLowerInvariant()}}}"" behavior=""0"" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
";
    }

    private static string BuildCustomizationsXml(VacationIdentity identity, bool includeFlowMgmtConnRef)
    {
        var startWf = BuildWorkflowElement(identity.StartWorkflowId, identity.StartFlowDisplayName, identity.StartWorkflowFileName);
        var endWf   = BuildWorkflowElement(identity.EndWorkflowId, identity.EndFlowDisplayName, identity.EndWorkflowFileName);

        var connRefs = new StringBuilder();
        connRefs.AppendLine($@"    <connectionreference connectionreferencelogicalname=""{identity.OutlookConnRefLogicalName}"">
      <connectionreferencedisplayname>{XmlEscape(OutlookConnReferenceDisplayName)}</connectionreferencedisplayname>
      <connectorid>{OutlookConnectorId}</connectorid>
      <iscustomizable>1</iscustomizable>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>");

        if (includeFlowMgmtConnRef)
        {
            connRefs.AppendLine($@"    <connectionreference connectionreferencelogicalname=""{identity.FlowMgmtConnRefLogicalName}"">
      <connectionreferencedisplayname>{XmlEscape(FlowMgmtConnReferenceDisplayName)}</connectionreferencedisplayname>
      <connectorid>{FlowMgmtConnectorId}</connectorid>
      <iscustomizable>1</iscustomizable>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>");
        }

        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<ImportExportXml xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <Entities />
  <Roles />
  <Workflows>
{startWf}
{endWf}
  </Workflows>
  <FieldSecurityProfiles />
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls />
  <EntityDataProviders />
  <connectionreferences>
{connRefs}  </connectionreferences>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>
";
    }

    private static string BuildWorkflowElement(Guid workflowId, string displayName, string fileName) =>
        $@"    <Workflow WorkflowId=""{{{workflowId.ToString("D").ToLowerInvariant()}}}"" Name=""{XmlEscape(displayName)}"">
      <JsonFileName>/Workflows/{XmlEscape(fileName)}</JsonFileName>
      <Type>1</Type>
      <Subprocess>0</Subprocess>
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>0</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <StateCode>1</StateCode>
      <StatusCode>2</StatusCode>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>1.0.0.0</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <BusinessProcessType>0</BusinessProcessType>
      <IsCustomProcessingStepAllowedForOtherPublishers>1</IsCustomProcessingStepAllowedForOtherPublishers>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode=""1033"" description=""{XmlEscape(displayName)}"" />
      </LocalizedNames>
    </Workflow>";

    private static string BuildContentTypesXml() => @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""xml"" ContentType=""application/octet-stream"" />
  <Default Extension=""json"" ContentType=""application/octet-stream"" />
</Types>
";

    private static string BuildReadme(string userEmail, VacationIdentity identity, DateTime start, DateTime end, bool hasScheduleFlowTarget)
    {
        var lines = new List<string>
        {
            "OofManager — Manual Vacation Solution package",
            "=============================================",
            "",
            "What this is",
            "------------",
            "A Dataverse solution containing two one-shot cloud flows that drive",
            "a vacation entirely from Microsoft 365:",
            "",
            $"  Start flow: '{identity.StartFlowDisplayName}'",
            $"     Fires once at {start:yyyy-MM-dd HH:mm} (local) — turns Outlook",
            "     automatic replies on and " + (hasScheduleFlowTarget
                ? "pauses your OofManager Weekly Schedule flow."
                : "(no Schedule-flow target was provided; only AutoReply is set)."),
            "",
            $"  End flow:   '{identity.EndFlowDisplayName}'",
            $"     Fires once at {end:yyyy-MM-dd HH:mm} (local) — " + (hasScheduleFlowTarget
                ? "re-enables your OofManager Weekly Schedule flow."
                : "no-op (Outlook clears AutoReply on its own at this time)."),
            "",
            "How to import",
            "-------------",
            "1. Open https://make.powerautomate.com and sign in as " + userEmail + ".",
            "2. Pick the environment named after you in the top-right env selector.",
            "3. Solutions → Import solution → Browse → pick this .zip.",
            "4. Map the 'OofManager Outlook' connection reference to your Office 365",
            "   Outlook connection (same connection the Weekly Schedule flow uses).",
            hasScheduleFlowTarget
                ? "5. Map the 'OofManager Flow Management' connection reference to a Power"
                : "5. (No Flow Management connection reference in this build.)",
            hasScheduleFlowTarget
                ? "   Automate Management connection (the importer will offer '+ New connection')."
                : "",
            "6. Click Import. Both flows are imported in the OFF state by default.",
            "7. Open each flow and click 'Turn on' so they actually fire.",
            "",
            "If you change your mind",
            "-----------------------",
            "Click 'Cancel planned vacation' in OofManager — it deletes this whole",
            $"solution ('{identity.SolutionUniqueName}'), which removes both flows in",
            "a single call.",
        };
        return string.Join("\r\n", lines.Where(l => l != null));
    }

    private static string XmlEscape(string s) => System.Security.SecurityElement.Escape(s) ?? string.Empty;

    private static VacationIdentity BuildIdentity(string userEmail)
    {
        var alias = SanitizeAlias(userEmail);

        var solutionUniqueName = $"OofManagerManualVacation_{alias}";
        var solutionDisplayName = $"OofManager Manual Vacation ({alias})";
        var startFlowDisplayName = $"OofManager Vacation Start ({alias})";
        var endFlowDisplayName   = $"OofManager Vacation End ({alias})";

        // Share the Outlook connection-reference logical name with
        // CloudSchedulePackageGenerator so users who already imported Cloud
        // Schedule get their existing Outlook connection reused
        // automatically by Dataverse (logical names identify rows; matching
        // name = same row = same bound connection). Without this, the
        // Vacation solution lands with an unbound Outlook ref and Enable-Flow
        // refuses to turn the flows on even though the connector type is
        // identical.
        var outlookConnRefLogical = $"{PublisherPrefix}_OofManagerOutlookConn_{alias.ToLowerInvariant()}";
        var flowMgmtConnRefLogical = $"{PublisherPrefix}_OofManagerVacFlowMgmtConn_{alias.ToLowerInvariant()}";

        var startWorkflowId = DeterministicGuid(VacationStartWorkflowNamespace, alias);
        var endWorkflowId   = DeterministicGuid(VacationEndWorkflowNamespace, alias);
        var outlookConnRefId = DeterministicGuid(OutlookConnRefNamespace, alias);
        var flowMgmtConnRefId = DeterministicGuid(FlowMgmtConnRefNamespace, alias);

        return new VacationIdentity(
            alias: alias,
            solutionUniqueName: solutionUniqueName,
            solutionDisplayName: solutionDisplayName,
            startFlowDisplayName: startFlowDisplayName,
            endFlowDisplayName: endFlowDisplayName,
            startWorkflowFileName: BuildWorkflowFileName(startFlowDisplayName, startWorkflowId),
            endWorkflowFileName: BuildWorkflowFileName(endFlowDisplayName, endWorkflowId),
            startWorkflowId: startWorkflowId,
            endWorkflowId: endWorkflowId,
            outlookConnRefId: outlookConnRefId,
            flowMgmtConnRefId: flowMgmtConnRefId,
            outlookConnRefLogicalName: outlookConnRefLogical,
            flowMgmtConnRefLogicalName: flowMgmtConnRefLogical);
    }

    private static string SanitizeAlias(string emailOrAlias)
    {
        var raw = emailOrAlias ?? string.Empty;
        var at = raw.IndexOf('@');
        var local = at >= 0 ? raw.Substring(0, at) : raw;

        var sb = new StringBuilder(local.Length);
        foreach (var c in local)
        {
            if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9'))
                sb.Append(c);
        }
        var cleaned = sb.ToString();
        if (cleaned.Length == 0) return "user";
        if (cleaned[0] >= '0' && cleaned[0] <= '9') cleaned = "u" + cleaned;
        return cleaned;
    }

    private static Guid DeterministicGuid(Guid namespaceId, string name)
    {
        var nsBytes = namespaceId.ToByteArray();
        SwapGuidByteOrder(nsBytes);

        var nameBytes = Encoding.UTF8.GetBytes(name);
        var input = new byte[nsBytes.Length + nameBytes.Length];
        Buffer.BlockCopy(nsBytes, 0, input, 0, nsBytes.Length);
        Buffer.BlockCopy(nameBytes, 0, input, nsBytes.Length, nameBytes.Length);

        byte[] hash;
        using (var sha1 = System.Security.Cryptography.SHA1.Create())
        {
            hash = sha1.ComputeHash(input);
        }

        var newGuid = new byte[16];
        Array.Copy(hash, newGuid, 16);
        newGuid[6] = (byte)((newGuid[6] & 0x0F) | (5 << 4));
        newGuid[8] = (byte)((newGuid[8] & 0x3F) | 0x80);

        SwapGuidByteOrder(newGuid);
        return new Guid(newGuid);
    }

    private static void SwapGuidByteOrder(byte[] guid)
    {
        (guid[0], guid[3]) = (guid[3], guid[0]);
        (guid[1], guid[2]) = (guid[2], guid[1]);
        (guid[4], guid[5]) = (guid[5], guid[4]);
        (guid[6], guid[7]) = (guid[7], guid[6]);
    }

    private static string PlainTextToHtml(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var encoded = WebUtility.HtmlEncode(value.Trim());
        encoded = encoded.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "<br>");
        return $"<html><body>{encoded}</body></html>";
    }

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };
}
