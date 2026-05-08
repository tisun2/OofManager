using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace OofManager.Wpf.Services;

/// <summary>
/// Generates a Dataverse <em>Solution</em> .zip that, when uploaded under
/// <em>Power Automate &rarr; Solutions &rarr; Import solution</em>, creates a
/// scheduled cloud flow which mirrors the user's local OofManager schedule.
/// The flow runs in Microsoft 365 itself, so the next OOF window keeps being
/// scheduled in Outlook even when every local computer is off.
///
/// <para>
/// We use the Dataverse solution format (rather than the older
/// "Import Package (Legacy)" format) because most managed tenants &mdash;
/// notably anything inside microsoft.com &mdash; have the
/// <em>"Create in Dataverse solutions"</em> admin policy enabled, which
/// disables the legacy import path entirely. Solution import is the only
/// path that works under that policy.
/// </para>
///
/// The zip contains:
/// <code>
///   [Content_Types].xml                                 &mdash; MIME registry required by every solution zip
///   solution.xml                                        &mdash; UniqueName, version, Publisher
///   customizations.xml                                  &mdash; lists the workflow as a solution component
///   Workflows/{Name}-{Guid}.json                        &mdash; the flow definition
///   OofManager-README.txt                               &mdash; human-readable import instructions
/// </code>
/// </summary>
public static class CloudSyncPackageGenerator
{
    // Stable publisher + solution identity. Reusing the same GUIDs across
    // generations means a re-import upgrades the existing solution rather
    // than creating "OofManagerCloudSync_2", "OofManagerCloudSync_3", etc.
    private const string PublisherUniqueName = "OofManagerPublisher";
    private const string PublisherDisplayName = "OofManager";
    private const string PublisherPrefix = "ofm";
    private const int PublisherCustomizationOption = 10000;
    // Stable workflow id — the flow's identity inside the solution. Keeping
    // this stable means re-importing produces an UPDATE to the existing
    // flow instead of duplicating it.
    private static readonly Guid WorkflowId = new("d2e91a8f-4b21-4d72-9c54-1a3b5c7e9f01");
    // Stable connection reference id — keeping this fixed means a re-import
    // upgrades the same connection reference rather than creating duplicates.
    private static readonly Guid ConnectionReferenceId = new("5c1d8b2a-9f63-4d8a-bc7e-1f3e6a9c2d05");
    private const string ConnectionReferenceLogicalName = "ofm_OofManagerOutlookConn";
    private const string ConnectionReferenceDisplayName = "OofManager Outlook";
    private const string ConnectorId = "/providers/Microsoft.PowerApps/apis/shared_office365";
    private const string SolutionUniqueName = "OofManagerCloudSync";
    private const string SolutionDisplayName = "OofManager Cloud Sync";
    private const string SolutionVersion = "1.0.0.0";
    private const string FlowDisplayName = "OofManager Cloud Sync";

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

        var tzId = TimeZoneInfo.Local.Id;
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

        var workflowJson = BuildWorkflowFileJson(
            workflowName: FlowDisplayName,
            tzId: tzId,
            triggerHour: triggerEnd.Hours,
            triggerMinute: triggerEnd.Minutes,
            triggerDays: triggerDays,
            startExpr: startExpr,
            endExpr: endExpr,
            audience: externalAudienceAll ? "all" : "contactsOnly",
            internalReply: internalReply,
            externalReply: externalReply);

        var solutionXml = BuildSolutionXml();
        var customizationsXml = BuildCustomizationsXml();
        var contentTypesXml = BuildContentTypesXml();

        if (File.Exists(outputPath))
        {
            try { File.Delete(outputPath); } catch { /* best-effort */ }
        }
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        using (var fs = File.Create(outputPath))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // Forward slashes only (backslashes get rejected by the importer).
            WriteEntry(zip, "[Content_Types].xml", contentTypesXml);
            WriteEntry(zip, "solution.xml", solutionXml);
            WriteEntry(zip, "customizations.xml", customizationsXml);
            WriteEntry(zip, $"Workflows/{WorkflowFileName}", workflowJson);
            WriteEntry(zip, "OofManager-README.txt", BuildReadme(userEmail));
        }

        return outputPath;
    }

    private static string WorkflowFileName =>
        // Convention used by the Power Platform exporter: display name with
        // spaces replaced by hyphens, followed by the workflow guid in braces,
        // .json. The actual file name doesn't have to match exactly, but the
        // <JsonFileName> reference in customizations.xml does have to point
        // at this same name.
        $"{FlowDisplayName.Replace(' ', '_')}-{WorkflowId.ToString("D").ToUpperInvariant()}.json";

    private static void WriteEntry(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    /// <summary>
    /// Workflow file inside <c>/Workflows/</c>. Mirrors the shape produced by
    /// <em>Solutions &rarr; Export</em> for a cloud flow: a <c>properties</c>
    /// envelope containing <c>connectionReferences</c>, the inline Logic Apps
    /// <c>definition</c>, and a flow display name.
    /// </summary>
    private static string BuildWorkflowFileJson(
        string workflowName,
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
        var wrapper = new Dictionary<string, object?>
        {
            ["properties"] = new Dictionary<string, object?>
            {
                ["connectionReferences"] = new Dictionary<string, object?>
                {
                    // Dataverse-resident flows resolve connectors via a
                    // ConnectionReference component instead of the embedded
                    // Connection map the legacy package used. The importer
                    // either matches an existing reference with the same
                    // logical name or prompts the user to create one.
                    ["shared_office365"] = new Dictionary<string, object?>
                    {
                        ["runtimeSource"] = "embedded",
                        ["connection"] = new Dictionary<string, object?>
                        {
                            ["connectionReferenceLogicalName"] = ConnectionReferenceLogicalName,
                        },
                        ["api"] = new Dictionary<string, object?>
                        {
                            ["name"] = "shared_office365",
                        },
                    },
                },
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
                            ["recurrence"] = BuildRecurrenceTrigger(triggerHour, triggerMinute, triggerDays, tzId),
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
                                    // The operationId in the Office 365 Outlook connector swagger is
                                    // SetAutomaticRepliesSetting_V2 (underscore before V2). Without the
                                    // underscore, the Power Automate designer can't resolve the action
                                    // and shows: "Unable to initialize operation details ... 404".
                                    ["operationId"] = "SetAutomaticRepliesSetting_V2",
                                    ["apiId"] = "/providers/Microsoft.PowerApps/apis/shared_office365",
                                },
                                ["parameters"] = new Dictionary<string, object?>
                                {
                                    // SetAutomaticRepliesSetting_V2 swagger has a single body
                                    // parameter named "body" whose schema wraps everything in
                                    // an "automaticRepliesSetting" object (which references
                                    // AutomaticRepliesSettingClient_V2). Pass the body as a
                                    // single parameter with a fully-nested object value — the
                                    // flat slash-key forms (with or without an
                                    // "automaticRepliesSetting/" prefix) are rejected by the
                                    // Power Automate designer with "X is no longer present in
                                    // the operation schema" for every key.
                                    ["body"] = new Dictionary<string, object?>
                                    {
                                        ["automaticRepliesSetting"] = new Dictionary<string, object?>
                                        {
                                            ["status"] = "scheduled",
                                            ["externalAudience"] = audience,
                                            ["scheduledStartDateTime"] = new Dictionary<string, object?>
                                            {
                                                ["dateTime"] = startExpr,
                                                ["timeZone"] = tzId,
                                            },
                                            ["scheduledEndDateTime"] = new Dictionary<string, object?>
                                            {
                                                ["dateTime"] = endExpr,
                                                ["timeZone"] = tzId,
                                            },
                                            ["internalReplyMessage"] = internalReply ?? string.Empty,
                                            ["externalReplyMessage"] = externalReply ?? string.Empty,
                                        },
                                    },
                                },
                                ["authentication"] = "@parameters('$authentication')",
                            },
                        },
                    },
                },
                ["parameters"] = new Dictionary<string, object?>(),
                ["displayName"] = workflowName,
            },
            ["schemaVersion"] = "1.0.0.0",
        };

        return JsonSerializer.Serialize(wrapper, JsonOpts);
    }

    private static Dictionary<string, object?> BuildRecurrenceTrigger(int triggerHour, int triggerMinute, string[] triggerDays, string tzId)
    {
        var hasWeekDays = triggerDays != null && triggerDays.Length > 0 && !string.IsNullOrWhiteSpace(triggerDays[0]);
        var schedule = new Dictionary<string, object?>
        {
            ["hours"] = new[] { triggerHour.ToString() },
            ["minutes"] = new[] { triggerMinute },
        };

        if (hasWeekDays)
        {
            schedule["weekDays"] = triggerDays;
        }

        return new Dictionary<string, object?>
        {
            ["frequency"] = hasWeekDays ? "Week" : "Day",
            ["interval"] = 1,
            ["schedule"] = schedule,
            ["timeZone"] = tzId,
        };
    }

    private static string BuildSolutionXml()
    {
        // Minimal solution manifest. UniqueName + Publisher prefix identify
        // the solution; the version triggers an upgrade vs. install. Component
        // type 29 = Workflow (covers both classic + cloud flows).
        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<ImportExportXml version=""9.2.0.1234"" SolutionPackageVersion=""9.2"" languagecode=""1033"" generatedBy=""OofManager"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <SolutionManifest>
    <UniqueName>{SolutionUniqueName}</UniqueName>
    <LocalizedNames>
      <LocalizedName description=""{XmlEscape(SolutionDisplayName)}"" languagecode=""1033"" />
    </LocalizedNames>
    <Descriptions>
      <Description description=""Schedules an OOF reply window every workday so Outlook stays in sync without your local computers being on. Generated by OofManager."" languagecode=""1033"" />
    </Descriptions>
    <Version>{SolutionVersion}</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>{PublisherUniqueName}</UniqueName>
      <LocalizedNames>
        <LocalizedName description=""{XmlEscape(PublisherDisplayName)}"" languagecode=""1033"" />
      </LocalizedNames>
      <Descriptions>
        <Description description=""Default publisher for OofManager-generated cloud sync solutions."" languagecode=""1033"" />
      </Descriptions>
      <EMailAddress xsi:nil=""true""></EMailAddress>
      <SupportingWebsiteUrl xsi:nil=""true""></SupportingWebsiteUrl>
      <CustomizationPrefix>{PublisherPrefix}</CustomizationPrefix>
      <CustomizationOptionValuePrefix>{PublisherCustomizationOption}</CustomizationOptionValuePrefix>
      <Addresses>
        <Address>
          <AddressNumber>1</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          <City xsi:nil=""true""></City>
          <County xsi:nil=""true""></County>
          <Country xsi:nil=""true""></Country>
          <Fax xsi:nil=""true""></Fax>
          <FreightTermsCode xsi:nil=""true""></FreightTermsCode>
          <ImportSequenceNumber xsi:nil=""true""></ImportSequenceNumber>
          <Latitude xsi:nil=""true""></Latitude>
          <Line1 xsi:nil=""true""></Line1>
          <Line2 xsi:nil=""true""></Line2>
          <Line3 xsi:nil=""true""></Line3>
          <Longitude xsi:nil=""true""></Longitude>
          <Name xsi:nil=""true""></Name>
          <PostalCode xsi:nil=""true""></PostalCode>
          <PostOfficeBox xsi:nil=""true""></PostOfficeBox>
          <PrimaryContactName xsi:nil=""true""></PrimaryContactName>
          <ShippingMethodCode>1</ShippingMethodCode>
          <StateOrProvince xsi:nil=""true""></StateOrProvince>
          <Telephone1 xsi:nil=""true""></Telephone1>
          <Telephone2 xsi:nil=""true""></Telephone2>
          <Telephone3 xsi:nil=""true""></Telephone3>
          <TimeZoneRuleVersionNumber xsi:nil=""true""></TimeZoneRuleVersionNumber>
          <UPSZone xsi:nil=""true""></UPSZone>
          <UTCOffset xsi:nil=""true""></UTCOffset>
          <UTCConversionTimeZoneCode xsi:nil=""true""></UTCConversionTimeZoneCode>
        </Address>
        <Address>
          <AddressNumber>2</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          <City xsi:nil=""true""></City>
          <County xsi:nil=""true""></County>
          <Country xsi:nil=""true""></Country>
          <Fax xsi:nil=""true""></Fax>
          <FreightTermsCode xsi:nil=""true""></FreightTermsCode>
          <ImportSequenceNumber xsi:nil=""true""></ImportSequenceNumber>
          <Latitude xsi:nil=""true""></Latitude>
          <Line1 xsi:nil=""true""></Line1>
          <Line2 xsi:nil=""true""></Line2>
          <Line3 xsi:nil=""true""></Line3>
          <Longitude xsi:nil=""true""></Longitude>
          <Name xsi:nil=""true""></Name>
          <PostalCode xsi:nil=""true""></PostalCode>
          <PostOfficeBox xsi:nil=""true""></PostOfficeBox>
          <PrimaryContactName xsi:nil=""true""></PrimaryContactName>
          <ShippingMethodCode>1</ShippingMethodCode>
          <StateOrProvince xsi:nil=""true""></StateOrProvince>
          <Telephone1 xsi:nil=""true""></Telephone1>
          <Telephone2 xsi:nil=""true""></Telephone2>
          <Telephone3 xsi:nil=""true""></Telephone3>
          <TimeZoneRuleVersionNumber xsi:nil=""true""></TimeZoneRuleVersionNumber>
          <UPSZone xsi:nil=""true""></UPSZone>
          <UTCOffset xsi:nil=""true""></UTCOffset>
          <UTCConversionTimeZoneCode xsi:nil=""true""></UTCConversionTimeZoneCode>
        </Address>
      </Addresses>
    </Publisher>
    <RootComponents>
      <RootComponent type=""29"" id=""{{{WorkflowId.ToString("D").ToLowerInvariant()}}}"" behavior=""0"" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
";
    }

    private static string BuildCustomizationsXml()
    {
        // <Workflow> Category=5 = Modern (cloud) Flow. <JsonFileName> path
        // points at the JSON file under /Workflows/. <PrimaryEntity>none</...>
        // is the magic value for non-entity-bound flows; required for the
        // importer to skip looking up an associated table.
        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<ImportExportXml xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <Entities />
  <Roles />
  <Workflows>
    <Workflow WorkflowId=""{{{WorkflowId.ToString("D").ToLowerInvariant()}}}"" Name=""{XmlEscape(FlowDisplayName)}"">
      <JsonFileName>/Workflows/{XmlEscape(WorkflowFileName)}</JsonFileName>
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
        <LocalizedName languagecode=""1033"" description=""{XmlEscape(FlowDisplayName)}"" />
      </LocalizedNames>
    </Workflow>
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
    <connectionreference connectionreferencelogicalname=""{ConnectionReferenceLogicalName}"">
      <connectionreferencedisplayname>{XmlEscape(ConnectionReferenceDisplayName)}</connectionreferencedisplayname>
      <connectorid>{ConnectorId}</connectorid>
      <iscustomizable>1</iscustomizable>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>
  </connectionreferences>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>
";
    }

    private static string BuildContentTypesXml()
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""xml"" ContentType=""application/octet-stream"" />
  <Default Extension=""json"" ContentType=""application/octet-stream"" />
</Types>
";
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
            "OofManager — Power Automate Solution package",
            "============================================",
            "",
            "What this is",
            "------------",
            "A Dataverse solution containing a single scheduled cloud flow that pushes",
            "your next out-of-office window to Outlook every workday. The flow runs",
            "entirely inside Microsoft 365, so it keeps working even when all of your",
            "computers are powered off.",
            "",
            "How to import",
            "-------------",
            "1. Open https://make.powerautomate.com and sign in as " + userEmail + ".",
            "2. Make sure the environment selector (top right) shows your default",
            "   environment — solutions are scoped to a single environment.",
            "3. In the left rail, click 'Solutions'.",
            "4. Click 'Import solution' on the toolbar.",
            "5. Click 'Browse', pick the .zip you got from OofManager, then 'Next'.",
            "6. The wizard shows the solution details. Click 'Next'.",
            "7. Connection References: the importer asks you to map the",
            "   'OofManager Outlook' connection reference to an Office 365 Outlook",
            "   connection. Either pick an existing connection or click",
            "   '+ New connection' and authorize one. Then 'Next'.",
            "8. Click 'Import'. Wait for 'Solution imported successfully' (~30s).",
            "9. Open the solution, click the 'OofManager Cloud Sync' flow, and",
            "   click 'Turn on' in the toolbar (imported solution flows are off",
            "   by default for safety).",
            "10. Click 'Test' -> 'Manually' -> 'Test' to fire it once and verify",
            "    Outlook accepts the schedule.",
            "",
            "If anything fails",
            "-----------------",
            "Re-run 'Set up cloud sync' in OofManager — the regenerated zip uses",
            "the same solution + workflow GUIDs, so a re-import becomes an upgrade",
            "rather than a duplicate. If solution import is also blocked by your",
            "tenant, switch to the manual setup guide (the other button in the",
            "OofManager UI).",
        });
    }

    private static string XmlEscape(string s) => System.Security.SecurityElement.Escape(s) ?? string.Empty;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };
}
