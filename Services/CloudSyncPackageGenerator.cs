using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace OofManager.Wpf.Services;

/// <summary>
/// Generates a Power Automate "Import Package (Legacy)" .zip that, when
/// uploaded under <em>My flows &rarr; Import &rarr; Import Package</em>,
/// creates a scheduled cloud flow which mirrors the user's local OofManager
/// schedule. The flow runs in Microsoft 365 itself, so the next OOF window
/// keeps being scheduled in Outlook even when every local computer is off.
///
/// The zip contains four files:
///   <code>manifest.json</code>                                 &mdash; package-level metadata + resource map
///   <code>Microsoft.Flow/flows/{guid}/definition.json</code>   &mdash; the actual flow definition
///   <code>Microsoft.Flow/flows/{guid}/apisMap.json</code>      &mdash; apiId &rarr; resource-uuid map
///   <code>Microsoft.Flow/flows/{guid}/connectionsMap.json</code> &mdash; connection-name &rarr; resource-uuid map
///
/// Schema is reverse-engineered from packages exported via Power Automate's
/// own UI; Microsoft does not publish a formal spec, but the format has been
/// stable for several years.
/// </summary>
public static class CloudSyncPackageGenerator
{
    /// <summary>
    /// Builds the package and writes it to <paramref name="outputPath"/>.
    /// Pass null to drop it under <c>%TEMP%\OofManager-CloudSync.zip</c>.
    /// Returns the resolved path so the caller can surface it in the UI.
    /// </summary>
    public static string Generate(
        WorkScheduleSnapshot schedule,
        string userEmail,
        string internalReply,
        string externalReply,
        bool externalAudienceAll,
        string? outputPath = null)
    {
        outputPath ??= Path.Combine(Path.GetTempPath(), "OofManager-CloudSync.zip");

        // Stable resource ids inside this package. They never leave the zip,
        // so reusing the same constants every run keeps re-imports as updates
        // rather than fresh installs (Power Automate keys off these GUIDs).
        // The flow id IS visible in Power Automate, so we still let the
        // import wizard rename the actual flow if a duplicate exists.
        var flowResId = "DEAD0001-0000-0000-0000-OOFMGRCLOUDSY".ToLowerInvariant().Replace("oofmgrcloudsy", "00f0000f10ce");
        // Fall back to a random-but-deterministic GUID if the cute string
        // above ever fails the GUID parser on some .NET version.
        if (!Guid.TryParse(flowResId, out _)) flowResId = Guid.NewGuid().ToString("D");

        var connResId = Guid.NewGuid().ToString("D");
        var apiResId = Guid.NewGuid().ToString("D");
        var flowFolder = Guid.NewGuid().ToString("D"); // GUID under Microsoft.Flow/flows/
        var packageTelemetryId = Guid.NewGuid().ToString("D");

        var tzId = TimeZoneInfo.Local.Id;

        // Trigger schedule: fire on the user's chosen workdays at the
        // earliest configured end-of-work time. Per-day variation is
        // acknowledged via the End Time expression below; the trigger
        // itself just needs to fire once per workday after work ends.
        var triggerEnd = ComputeRepresentativeEnd(schedule);
        var workStart = ComputeRepresentativeStart(schedule);

        var weekDays = new[]
        {
            DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
            DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday,
        };
        var triggerDays = weekDays.Where(schedule.IsWorkday)
            .Select(d => d.ToString())
            .ToArray();

        // Per-day-of-week hop count (today → next configured workday).
        // Used by the End-time expression so e.g. Friday OOF runs through
        // Monday morning on a Mon–Fri schedule.
        var hopDays = new int[7];
        for (int i = 0; i < 7; i++)
        {
            for (int hop = 1; hop <= 7; hop++)
            {
                var candidate = (DayOfWeek)(((int)(DayOfWeek)i + hop) % 7);
                if (schedule.IsWorkday(candidate)) { hopDays[i] = hop; break; }
            }
        }

        var startExpr = $"@{{formatDateTime(convertFromUtc(utcNow(), '{tzId}'), 'yyyy-MM-ddTHH:mm:ss')}}";
        var endExpr = $"@{{{BuildEndTimeExpression(hopDays, workStart.Hours, workStart.Minutes, tzId)}}}";

        var definitionJson = BuildDefinitionJson(
            flowFolder: flowFolder,
            displayName: "OofManager Cloud Sync",
            tzId: tzId,
            triggerHour: triggerEnd.Hours,
            triggerMinute: triggerEnd.Minutes,
            triggerDays: triggerDays,
            startExpr: startExpr,
            endExpr: endExpr,
            audience: externalAudienceAll ? "all" : "contactsOnly",
            internalReply: internalReply,
            externalReply: externalReply);

        var manifestJson = BuildManifestJson(
            flowResId: flowResId,
            connResId: connResId,
            apiResId: apiResId,
            packageTelemetryId: packageTelemetryId,
            flowDisplayName: "OofManager Cloud Sync",
            creator: userEmail);

        var apisMap = "{\"shared_office365\":\"" + apiResId + "\"}";
        var connsMap = "{\"shared_office365\":\"" + connResId + "\"}";

        // Always overwrite. The package is regenerated every click, so a
        // stale zip from yesterday's schedule would be misleading.
        if (File.Exists(outputPath))
        {
            try { File.Delete(outputPath); } catch { /* best-effort */ }
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // ZipArchive auto-creates parent folders inside the archive when the
        // entry name contains slashes. Power Automate's import is sensitive
        // to forward slashes (Windows backslashes get rejected as "invalid
        // package layout"), so build the entry names explicitly.
        using (var fs = File.Create(outputPath))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            WriteEntry(zip, "manifest.json", manifestJson);
            WriteEntry(zip, $"Microsoft.Flow/flows/{flowFolder}/definition.json", definitionJson);
            WriteEntry(zip, $"Microsoft.Flow/flows/{flowFolder}/apisMap.json", apisMap);
            WriteEntry(zip, $"Microsoft.Flow/flows/{flowFolder}/connectionsMap.json", connsMap);
            WriteEntry(zip, "OofManager-README.txt", BuildReadme(userEmail));
        }

        return outputPath;
    }

    private static void WriteEntry(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    private static string BuildDefinitionJson(
        string flowFolder,
        string displayName,
        string tzId,
        int triggerHour,
        int triggerMinute,
        string[] triggerDays,
        string startExpr,
        string endExpr,
        string audience,
        string internalReply,
        string externalReply)
    {
        // Use Dictionary<string,object> directly so we can write the literal
        // "$schema", "$connections", "$authentication" keys without resorting
        // to post-process string replacement (which previously also clobbered
        // legitimate keys like "schema" in the manifest and "authentication"
        // inside inputs).
        var defObj = new Dictionary<string, object?>
        {
            ["name"] = flowFolder,
            ["id"] = $"/providers/Microsoft.Flow/flows/{flowFolder}",
            ["type"] = "Microsoft.Flow/flows",
            ["properties"] = new Dictionary<string, object?>
            {
                ["apiId"] = "/providers/Microsoft.PowerApps/apis/shared_logicflows",
                ["displayName"] = displayName,
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
                        ["Recurrence"] = new Dictionary<string, object?>
                        {
                            ["recurrence"] = new Dictionary<string, object?>
                            {
                                ["frequency"] = "Day",
                                ["interval"] = 1,
                                ["schedule"] = new Dictionary<string, object?>
                                {
                                    ["hours"] = new[] { triggerHour.ToString() },
                                    ["minutes"] = new[] { triggerMinute },
                                    ["weekDays"] = triggerDays,
                                },
                                ["timeZone"] = tzId,
                            },
                            ["type"] = "Recurrence",
                        },
                    },
                    ["actions"] = new Dictionary<string, object?>
                    {
                        ["Set up automatic replies (V2)"] = new Dictionary<string, object?>
                        {
                            ["runAfter"] = new Dictionary<string, object?>(),
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
                                    ["operationId"] = "SetAutomaticRepliesSettingV2",
                                    ["apiId"] = "/providers/Microsoft.PowerApps/apis/shared_office365",
                                },
                                ["parameters"] = new Dictionary<string, object?>
                                {
                                    ["automaticRepliesSetting/status"] = "scheduled",
                                    ["automaticRepliesSetting/externalAudience"] = audience,
                                    ["automaticRepliesSetting/scheduledStartDateTime/dateTime"] = startExpr,
                                    ["automaticRepliesSetting/scheduledStartDateTime/timeZone"] = tzId,
                                    ["automaticRepliesSetting/scheduledEndDateTime/dateTime"] = endExpr,
                                    ["automaticRepliesSetting/scheduledEndDateTime/timeZone"] = tzId,
                                    ["automaticRepliesSetting/internalReplyMessage"] = internalReply ?? string.Empty,
                                    ["automaticRepliesSetting/externalReplyMessage"] = externalReply ?? string.Empty,
                                },
                                ["authentication"] = "@parameters('$authentication')",
                            },
                        },
                    },
                },
                ["connectionReferences"] = new Dictionary<string, object?>
                {
                    ["shared_office365"] = new Dictionary<string, object?>
                    {
                        ["connectionName"] = "shared-office365-placeholder",
                        ["source"] = "Embedded",
                        ["id"] = "/providers/Microsoft.PowerApps/apis/shared_office365",
                        ["tier"] = "NotSpecified",
                    },
                },
                ["flowFailureAlertSubscribed"] = false,
            },
            ["schemaVersion"] = "1.0.0.0",
        };

        return JsonSerializer.Serialize(defObj, JsonOpts);
    }

    private static string BuildManifestJson(
        string flowResId,
        string connResId,
        string apiResId,
        string packageTelemetryId,
        string flowDisplayName,
        string creator)
    {
        // Manifest top-level uses the literal property name "schema" (no
        // dollar sign) — only the inner Logic Apps definition uses "$schema".
        var manifest = new Dictionary<string, object?>
        {
            ["schema"] = "1.0",
            ["details"] = new Dictionary<string, object?>
            {
                ["displayName"] = flowDisplayName,
                ["description"] = "Generated by OofManager. Schedules an OOF reply window every workday so Outlook stays in sync without your local computers being on.",
                ["createdTime"] = DateTime.UtcNow.ToString("O"),
                ["packageTelemetryId"] = packageTelemetryId,
                ["creator"] = creator,
                ["sourceEnvironment"] = "default",
            },
            ["resources"] = new Dictionary<string, object?>
            {
                [flowResId] = new Dictionary<string, object?>
                {
                    ["id"] = null,
                    ["name"] = flowDisplayName,
                    ["type"] = "Microsoft.Flow/flows",
                    ["suggestedCreationType"] = "New",
                    ["creationType"] = "New, Update, Existing",
                    ["details"] = new Dictionary<string, object?> { ["displayName"] = flowDisplayName },
                    ["configurableBy"] = "User",
                    ["hierarchy"] = "Root",
                    ["dependsOn"] = new[] { connResId },
                },
                [connResId] = new Dictionary<string, object?>
                {
                    ["id"] = null,
                    ["name"] = "shared_office365",
                    ["type"] = "Microsoft.PowerApps/apis/connections",
                    ["suggestedCreationType"] = "Existing",
                    ["creationType"] = "Existing",
                    ["details"] = new Dictionary<string, object?> { ["displayName"] = "Office 365 Outlook" },
                    ["configurableBy"] = "User",
                    ["hierarchy"] = "Child",
                    ["dependsOn"] = new[] { apiResId },
                },
                [apiResId] = new Dictionary<string, object?>
                {
                    ["id"] = "/providers/Microsoft.PowerApps/apis/shared_office365",
                    ["name"] = "shared_office365",
                    ["type"] = "Microsoft.PowerApps/apis",
                    ["suggestedCreationType"] = "Existing",
                    ["creationType"] = "Existing",
                    ["details"] = new Dictionary<string, object?> { ["displayName"] = "Office 365 Outlook" },
                    ["configurableBy"] = "System",
                    ["hierarchy"] = "Child",
                    // No dependsOn — the connector is a leaf.
                },
            },
        };

        return JsonSerializer.Serialize(manifest, JsonOpts);
    }

    private static string BuildEndTimeExpression(int[] hopDays, int hour, int minute, string tzId)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < 6; i++)
            sb.Append("if(equals(dayOfWeek(utcNow()),").Append(i).Append("),").Append(hopDays[i]).Append(',');
        sb.Append(hopDays[6]);
        for (int i = 0; i < 6; i++) sb.Append(')');
        var hop = sb.ToString();

        var localToday = $"startOfDay(convertFromUtc(utcNow(), '{tzId}'))";
        return $"formatDateTime(addMinutes(addHours(addDays({localToday}, {hop}), {hour}), {minute}), 'yyyy-MM-ddTHH:mm:ss')";
    }

    private static TimeSpan ComputeRepresentativeEnd(WorkScheduleSnapshot s)
    {
        TimeSpan? earliest = null;
        foreach (DayOfWeek d in Enum.GetValues(typeof(DayOfWeek)))
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
        foreach (DayOfWeek d in Enum.GetValues(typeof(DayOfWeek)))
        {
            if (!s.IsWorkday(d)) continue;
            var v = s.GetStart(d);
            if (latest == null || v > latest.Value) latest = v;
        }
        return latest ?? new TimeSpan(9, 0, 0);
    }

    private static string BuildReadme(string userEmail)
    {
        return string.Join("\r\n", new[]
        {
            "OofManager — Power Automate import package",
            "===========================================",
            "",
            "What this is",
            "------------",
            "A scheduled cloud flow that pushes your next out-of-office window to",
            "Outlook every workday. It runs entirely inside Microsoft 365, so it",
            "keeps working even when all of your computers are powered off.",
            "",
            "How to import",
            "-------------",
            "1. Open https://make.powerautomate.com and sign in as " + userEmail + ".",
            "2. In the left rail, click 'My flows'.",
            "3. Click 'Import' (top toolbar) -> 'Import Package (Legacy)'.",
            "4. Click 'Upload' and pick the .zip you got from OofManager.",
            "5. Power Automate shows the package contents. Under 'Related resources',",
            "   click 'Select during import' next to the Office 365 Outlook",
            "   connection, then either pick an existing connection or click",
            "   '+ Create new' to authorize one.",
            "6. Click 'Import' at the bottom right. The flow appears in 'My flows'.",
            "7. Open the flow, click 'Edit' once to confirm everything looks right,",
            "   then 'Save'. (Power Automate sometimes needs a manual save to",
            "   activate imported flows.)",
            "8. Click 'Test' -> 'Manually' -> 'Test' to fire it once and verify",
            "   Outlook accepts the schedule.",
            "",
            "If anything fails",
            "-----------------",
            "Re-run 'Set up cloud sync' in OofManager — that regenerates this zip",
            "with the same flow id, so a re-import becomes an update rather than a",
            "fresh duplicate. If your tenant blocks legacy package import, switch",
            "to the manual setup guide (the other button in the OofManager UI).",
        });
    }

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        // Power Automate requires UTF-8 without BOM and tolerates any escaping
        // mode; default is fine. Ensure forward slashes in URLs aren't escaped
        // (the import validator gets confused by \/ in apiId values).
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };
}
