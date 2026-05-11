using System.IO;
using System.IO.Compression;
using System.Net;
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
    // Publisher identity is shared across every OofManager-generated solution
    // (Dataverse explicitly supports many solutions per publisher). Per-user
    // identifiers — solution unique name, workflow id, connection-reference
    // id — are derived from the signed-in user's mailbox alias instead, so
    // two people in the same tenant default environment don't clobber each
    // other's import. Same alias deterministically produces the same GUIDs
    // across re-imports, which preserves the "re-import = upgrade, not
    // duplicate" behavior we want for a single user.
    private const string PublisherUniqueName = "OofManagerPublisher";
    private const string PublisherDisplayName = "OofManager";
    private const string PublisherPrefix = "ofm";
    private const int PublisherCustomizationOption = 10000;
    // Namespaces for the deterministic v5 GUIDs. These were the previously
    // hard-coded WorkflowId / ConnectionReferenceId; reusing them as
    // namespaces keeps the value space disjoint from any other GUID we
    // might mint elsewhere in the app.
    private static readonly Guid WorkflowIdNamespace = new("d2e91a8f-4b21-4d72-9c54-1a3b5c7e9f01");
    private static readonly Guid ConnectionReferenceIdNamespace = new("5c1d8b2a-9f63-4d8a-bc7e-1f3e6a9c2d05");
    private const string ConnectionReferenceDisplayName = "OofManager Outlook";
    private const string ConnectorId = "/providers/Microsoft.PowerApps/apis/shared_office365";
    private const string SolutionVersion = "1.0.0.0";

    /// <summary>
    /// All per-user identifiers needed to stamp the solution package. Built
    /// once at the top of <see cref="Generate"/> from the signed-in user's
    /// alias and threaded into each Build* helper so the same name/GUID is
    /// used in solution.xml, customizations.xml, and the workflow JSON.
    /// </summary>
    private sealed class CloudSyncIdentity
    {
        public string Alias { get; }
        public string SolutionUniqueName { get; }
        public string SolutionDisplayName { get; }
        public string FlowDisplayName { get; }
        public string WorkflowFileName { get; }
        public Guid WorkflowId { get; }
        public Guid ConnectionReferenceId { get; }
        public string ConnectionReferenceLogicalName { get; }

        public CloudSyncIdentity(
            string alias,
            string solutionUniqueName,
            string solutionDisplayName,
            string flowDisplayName,
            string workflowFileName,
            Guid workflowId,
            Guid connectionReferenceId,
            string connectionReferenceLogicalName)
        {
            Alias = alias;
            SolutionUniqueName = solutionUniqueName;
            SolutionDisplayName = solutionDisplayName;
            FlowDisplayName = flowDisplayName;
            WorkflowFileName = workflowFileName;
            WorkflowId = workflowId;
            ConnectionReferenceId = connectionReferenceId;
            ConnectionReferenceLogicalName = connectionReferenceLogicalName;
        }
    }

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
        bool generateManaged = true,
        string? outputPath = null)
    {
        outputPath ??= Path.Combine(Path.GetTempPath(), "OofManager-CloudSync.zip");

        var identity = BuildIdentity(userEmail);

        var tzId = TimeZoneInfo.Local.Id;
        // Trigger fires at the earliest end-of-shift across all workdays.
        // Days with a later end (e.g. Fri 18:00 on a Mon-Thu 17:30 / Fri 18:00
        // schedule) get a future-dated OOF window that activates at their
        // actual end-of-shift, computed in the action body via per-dow lookup.
        var triggerEnd = ComputeRepresentativeEnd(schedule);

        var weekDays = new[]
        {
            DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
            DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday,
        };
        var triggerDays = weekDays.Where(schedule.IsWorkday)
            .Select(d => d.ToString())
            .ToArray();

        // Per-day-of-week lookup tables, indexed by Power Automate's
        // dayOfWeek (0=Sunday..6=Saturday — same as .NET DayOfWeek).
        // Today's actual end-of-shift hour:minute (used as OOF.start so days
        // with a later end-of-shift than the trigger get the correct start
        // timestamp), plus the hop-count and start-time of the next workday
        // (used as OOF.end so a Friday OOF runs through Monday morning even
        // when Mon's start-of-shift differs from Tue's).
        var todayEndH = new int[7];
        var todayEndM = new int[7];
        var hopDays = new int[7];
        var nextStartH = new int[7];
        var nextStartM = new int[7];
        for (int dow = 0; dow < 7; dow++)
        {
            var day = (DayOfWeek)dow;
            if (schedule.IsWorkday(day))
            {
                var end = schedule.GetEnd(day);
                todayEndH[dow] = end.Hours;
                todayEndM[dow] = end.Minutes;
            }
            // Non-workdays leave todayEnd at 0; the trigger's weekDays filter
            // means the action never fires on them anyway.
            for (int hop = 1; hop <= 7; hop++)
            {
                var candidate = (DayOfWeek)((dow + hop) % 7);
                if (schedule.IsWorkday(candidate))
                {
                    hopDays[dow] = hop;
                    var start = schedule.GetStart(candidate);
                    nextStartH[dow] = start.Hours;
                    nextStartM[dow] = start.Minutes;
                    break;
                }
            }
        }

        // Both expressions use the same shape:
        //   addDays(localToday, hopDays[dow]) + hour[dow] + minute[dow]
        // For start, hopDays is all-zeros (= today). For end, hopDays jumps
        // ahead to the next configured workday and we pick up that day's
        // start-of-shift.
        var hopZero = new int[7];
        var startExpr = $"@{{{BuildPerDowTimestampExpression(hopZero, todayEndH, todayEndM, tzId)}}}";
        var endExpr = $"@{{{BuildPerDowTimestampExpression(hopDays, nextStartH, nextStartM, tzId)}}}";

        var workflowJson = BuildWorkflowFileJson(
            identity: identity,
            tzId: tzId,
            triggerHour: triggerEnd.Hours,
            triggerMinute: triggerEnd.Minutes,
            triggerDays: triggerDays,
            startExpr: startExpr,
            endExpr: endExpr,
            audience: externalAudienceAll ? "all" : "contactsOnly",
            // Outlook renders the reply message field as HTML — plain newlines
            // get flattened into a single paragraph in the Automatic Replies
            // dialog. ExchangeService applies the same wrap on the local PS
            // path; mirror it here so cloud-pushed replies preserve line
            // breaks the same way the local sync does.
            internalReply: PlainTextToHtml(internalReply),
            externalReply: PlainTextToHtml(externalReply));

        var solutionXml = BuildSolutionXml(generateManaged, identity);
        var customizationsXml = BuildCustomizationsXml(identity);
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
            WriteEntry(zip, $"Workflows/{identity.WorkflowFileName}", workflowJson);
            WriteEntry(zip, "OofManager-README.txt", BuildReadme(userEmail, identity));
        }

        return outputPath;
    }

    private static string BuildWorkflowFileName(string flowDisplayName, Guid workflowId) =>
        // Convention used by the Power Platform exporter: display name with
        // spaces replaced by underscores, followed by the workflow guid in
        // braces, .json. The actual file name doesn't have to match exactly,
        // but the <JsonFileName> reference in customizations.xml does have
        // to point at this same name.
        $"{flowDisplayName.Replace(' ', '_')}-{workflowId.ToString("D").ToUpperInvariant()}.json";

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
        CloudSyncIdentity identity,
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
                            ["connectionReferenceLogicalName"] = identity.ConnectionReferenceLogicalName,
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
                                    // Logic Apps splits a nested-object body into individual
                                    // designer fields when the inputs.parameters keys use
                                    // slash-separated paths starting with the body parameter
                                    // name. Without the "body/" prefix the designer reports
                                    // every key as "no longer present in operation schema";
                                    // with a single nested "body" object the action runs but
                                    // every value is buried in the raw Body box.
                                    ["body/automaticRepliesSetting/status"] = "scheduled",
                                    ["body/automaticRepliesSetting/externalAudience"] = audience,
                                    ["body/automaticRepliesSetting/scheduledStartDateTime/dateTime"] = startExpr,
                                    ["body/automaticRepliesSetting/scheduledStartDateTime/timeZone"] = tzId,
                                    ["body/automaticRepliesSetting/scheduledEndDateTime/dateTime"] = endExpr,
                                    ["body/automaticRepliesSetting/scheduledEndDateTime/timeZone"] = tzId,
                                    ["body/automaticRepliesSetting/internalReplyMessage"] = internalReply ?? string.Empty,
                                    ["body/automaticRepliesSetting/externalReplyMessage"] = externalReply ?? string.Empty,
                                },
                                ["authentication"] = "@parameters('$authentication')",
                            },
                        },
                    },
                },
                ["parameters"] = new Dictionary<string, object?>(),
                ["displayName"] = identity.FlowDisplayName,
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

    private static string BuildSolutionXml(bool managed, CloudSyncIdentity identity)
    {
        // Minimal solution manifest. UniqueName + Publisher prefix identify
        // the solution; the version triggers an upgrade vs. install. Component
        // type 29 = Workflow (covers both classic + cloud flows).
        return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<ImportExportXml version=""9.2.0.1234"" SolutionPackageVersion=""9.2"" languagecode=""1033"" generatedBy=""OofManager"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <SolutionManifest>
    <UniqueName>{identity.SolutionUniqueName}</UniqueName>
    <LocalizedNames>
      <LocalizedName description=""{XmlEscape(identity.SolutionDisplayName)}"" languagecode=""1033"" />
    </LocalizedNames>
    <Descriptions>
      <Description description=""Schedules an OOF reply window every workday so Outlook stays in sync without your local computers being on. Generated by OofManager."" languagecode=""1033"" />
    </Descriptions>
    <Version>{SolutionVersion}</Version>
    <!-- Managed=1 so Power Automate honors StateCode=1/StatusCode=2 in
         customizations.xml and activates the flow automatically once the
         user finishes the import wizard's connection-binding step. With
         Managed=0 the flow always lands in the Off state and the user has
         to manually toggle it on, regardless of the StateCode values. -->
    <Managed>{(managed ? 1 : 0)}</Managed>
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
      <RootComponent type=""29"" id=""{{{identity.WorkflowId.ToString("D").ToLowerInvariant()}}}"" behavior=""0"" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
";
    }

    private static string BuildCustomizationsXml(CloudSyncIdentity identity)
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
    <Workflow WorkflowId=""{{{identity.WorkflowId.ToString("D").ToLowerInvariant()}}}"" Name=""{XmlEscape(identity.FlowDisplayName)}"">
      <JsonFileName>/Workflows/{XmlEscape(identity.WorkflowFileName)}</JsonFileName>
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
        <LocalizedName languagecode=""1033"" description=""{XmlEscape(identity.FlowDisplayName)}"" />
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
    <connectionreference connectionreferencelogicalname=""{identity.ConnectionReferenceLogicalName}"">
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

    private static string BuildPerDowTimestampExpression(int[] hopByDow, int[] hourByDow, int[] minuteByDow, string tzId)
    {
        // dayOfWeek operates on a string timestamp; feeding it the local-tz
        // string (vs. utcNow()) keeps the index correct around midnight in
        // negative-offset zones, and matches the day boundary the recurrence
        // trigger itself uses.
        var localDow = $"dayOfWeek(convertFromUtc(utcNow(), '{tzId}'))";
        var localToday = $"startOfDay(convertFromUtc(utcNow(), '{tzId}'))";
        var hopExpr = BuildPerDowLookup(hopByDow, localDow);
        var hourExpr = BuildPerDowLookup(hourByDow, localDow);
        var minuteExpr = BuildPerDowLookup(minuteByDow, localDow);
        return $"formatDateTime(addMinutes(addHours(addDays({localToday}, {hopExpr}), {hourExpr}), {minuteExpr}), 'yyyy-MM-ddTHH:mm:ss')";
    }

    private static string BuildPerDowLookup(int[] table, string dowExpr)
    {
        // Nested-if chain: if(dow=0, table[0], if(dow=1, table[1], ... table[6]))
        var sb = new StringBuilder();
        for (int i = 0; i < 6; i++)
            sb.Append("if(equals(").Append(dowExpr).Append(',').Append(i).Append("),").Append(table[i]).Append(',');
        sb.Append(table[6]);
        for (int i = 0; i < 6; i++) sb.Append(')');
        return sb.ToString();
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

    private static string BuildReadme(string userEmail, CloudSyncIdentity identity)
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
            "9. Open the solution, click the '" + identity.FlowDisplayName + "' flow,",
            "   and click 'Turn on' in the toolbar (imported solution flows are off",
            "   by default for safety).",
            "10. Click 'Test' -> 'Manually' -> 'Test' to fire it once and verify",
            "    Outlook accepts the schedule.",
            "",
            "If anything fails",
            "-----------------",
            "Re-run 'Set up cloud sync' in OofManager — the regenerated zip uses",
            "the same solution + workflow GUIDs (derived from your alias '" + identity.Alias + "'),",
            "so a re-import becomes an upgrade rather than a duplicate. If solution",
            "import is also blocked by your tenant, switch to the manual setup guide",
            "(the other button in the OofManager UI).",
        });
    }

    private static string XmlEscape(string s) => System.Security.SecurityElement.Escape(s) ?? string.Empty;

    /// <summary>
    /// Derives all per-user identifiers from the signed-in user's email
    /// address. Two users in the same tenant default environment will see
    /// distinct solution names + GUIDs (so neither clobbers the other on
    /// import), while the same user re-importing repeatedly always sees
    /// the same identifiers (so re-import = upgrade, not duplicate).
    /// </summary>
    private static CloudSyncIdentity BuildIdentity(string userEmail)
    {
        var alias = SanitizeAlias(userEmail);

        // Solution UniqueName must match [A-Za-z][A-Za-z0-9_]*. The base
        // is ASCII-safe; alias is already sanitized to alphanumerics by
        // SanitizeAlias, so concatenation stays valid.
        var solutionUniqueName = $"OofManagerCloudSync_{alias}";
        var solutionDisplayName = $"OofManager Cloud Sync ({alias})";
        var flowDisplayName = $"OofManager Cloud Sync ({alias})";

        // Connection-reference logical names must be prefixed with the
        // publisher's customization prefix and be Dataverse-safe. Lower
        // case the alias here because logical names are stored lower-case
        // by Dataverse anyway.
        var connRefLogicalName = $"{PublisherPrefix}_OofManagerOutlookConn_{alias.ToLowerInvariant()}";

        var workflowId = DeterministicGuid(WorkflowIdNamespace, alias);
        var connRefId = DeterministicGuid(ConnectionReferenceIdNamespace, alias);

        var workflowFileName = BuildWorkflowFileName(flowDisplayName, workflowId);

        return new CloudSyncIdentity(
            alias: alias,
            solutionUniqueName: solutionUniqueName,
            solutionDisplayName: solutionDisplayName,
            flowDisplayName: flowDisplayName,
            workflowFileName: workflowFileName,
            workflowId: workflowId,
            connectionReferenceId: connRefId,
            connectionReferenceLogicalName: connRefLogicalName);
    }

    /// <summary>
    /// Extracts the local part of an email address and strips it down to
    /// the characters Dataverse schema names allow (letters and digits).
    /// Falls back to "user" so the package still builds when the caller
    /// hasn't provided an email yet.
    /// </summary>
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
        // Schema names must start with a letter — prepend "u" if the alias
        // happens to start with a digit (rare but valid in some tenants).
        if (cleaned[0] >= '0' && cleaned[0] <= '9') cleaned = "u" + cleaned;
        return cleaned;
    }

    /// <summary>
    /// RFC 4122 §4.3 v5 (SHA-1) namespace UUID. Same input always produces
    /// the same GUID, so per-alias identifiers stay stable across
    /// re-generations of the package.
    /// </summary>
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
        // Set version (5) in the high nibble of byte 6.
        newGuid[6] = (byte)((newGuid[6] & 0x0F) | (5 << 4));
        // Set variant (RFC 4122) in the high bits of byte 8.
        newGuid[8] = (byte)((newGuid[8] & 0x3F) | 0x80);

        SwapGuidByteOrder(newGuid);
        return new Guid(newGuid);
    }

    private static void SwapGuidByteOrder(byte[] guid)
    {
        // .NET's Guid.ToByteArray serializes the first three fields in
        // little-endian order. RFC 4122 hashing operates on the network
        // (big-endian) layout, so we swap before hashing and after
        // assembling the result.
        (guid[0], guid[3]) = (guid[3], guid[0]);
        (guid[1], guid[2]) = (guid[2], guid[1]);
        (guid[4], guid[5]) = (guid[5], guid[4]);
        (guid[6], guid[7]) = (guid[7], guid[6]);
    }

    /// <summary>
    /// Wraps a plain-text reply in a minimal HTML envelope, the same shape
    /// ExchangeService produces on the local PowerShell path. Without this
    /// the Office 365 connector accepts the value but Outlook renders the
    /// reply as a single line because the field expects HTML markup.
    /// </summary>
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
