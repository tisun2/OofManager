namespace OofManager.Wpf.Services;

public enum PowerAutomateOutcome
{
    Success,
    NoFlowFound,
    SignInFailed,
    SolutionAwareBlocked,
    OtherError,
}
public sealed class PowerAutomateResult
{
    public PowerAutomateResult(
        PowerAutomateOutcome outcome,
        string message,
        IReadOnlyList<string> flowDisplayNames,
        IReadOnlyList<PowerAutomateFlowReference>? flowReferences = null)
    {
        Outcome = outcome;
        Message = message;
        FlowDisplayNames = flowDisplayNames;
        FlowReferences = flowReferences ?? Array.Empty<PowerAutomateFlowReference>();
    }

    public PowerAutomateOutcome Outcome { get; }
    public string Message { get; }
    public IReadOnlyList<string> FlowDisplayNames { get; }
    public IReadOnlyList<PowerAutomateFlowReference> FlowReferences { get; }
}

public sealed class PowerAutomateFlowReference
{
    public PowerAutomateFlowReference(string environmentName, string flowName, string displayName)
    {
        EnvironmentName = environmentName;
        FlowName = flowName;
        DisplayName = displayName;
    }

    public string EnvironmentName { get; }
    public string FlowName { get; }
    public string DisplayName { get; }
}

public enum PowerAutomateFlowState
{
    Unknown,
    On,
    Off,
    NotFound,
}

public sealed class PowerAutomateStatusResult
{
    public PowerAutomateStatusResult(
        PowerAutomateOutcome outcome,
        PowerAutomateFlowState state,
        string message,
        IReadOnlyList<string> flowDisplayNames,
        IReadOnlyList<PowerAutomateFlowReference>? flowReferences = null)
    {
        Outcome = outcome;
        State = state;
        Message = message;
        FlowDisplayNames = flowDisplayNames;
        FlowReferences = flowReferences ?? Array.Empty<PowerAutomateFlowReference>();
    }

    public PowerAutomateOutcome Outcome { get; }
    public PowerAutomateFlowState State { get; }
    public string Message { get; }
    public IReadOnlyList<string> FlowDisplayNames { get; }
    public IReadOnlyList<PowerAutomateFlowReference> FlowReferences { get; }
}

public enum CloudScheduleImportOutcome
{
    Success,
    SignInFailed,
    NoOwnedEnvironment,
    AlreadyExists,
    ImportFailed,
    TenantBlocked,
    OtherError,
}

/// <summary>
/// Outcome of <see cref="IPowerAutomateService.ImportCloudScheduleSolutionAsync"/>.
/// On <see cref="CloudScheduleImportOutcome.Success"/>, all four URL/id fields
/// are populated so callers can cache the imported flow reference and, when
/// possible, verify or turn on the flow after import. On <see cref="CloudScheduleImportOutcome.AlreadyExists"/>,
/// the same env metadata is populated so a confirmation dialog can name the
/// env about to be overwritten before the caller retries with <c>forceOverwrite=true</c>.
/// </summary>
public sealed class CloudScheduleImportResult
{
    public CloudScheduleImportResult(
        CloudScheduleImportOutcome outcome,
        string message,
        string? environmentId,
        string? environmentDisplayName,
        string? instanceUrl,
        string? workflowId)
    {
        Outcome = outcome;
        Message = message;
        EnvironmentId = environmentId;
        EnvironmentDisplayName = environmentDisplayName;
        InstanceUrl = instanceUrl;
        WorkflowId = workflowId;
    }

    public CloudScheduleImportOutcome Outcome { get; }
    public string Message { get; }
    public string? EnvironmentId { get; }
    public string? EnvironmentDisplayName { get; }
    public string? InstanceUrl { get; }
    public string? WorkflowId { get; }
}

/// <summary>
/// Snapshot of what's actually deployed in the cloud, decoded from the
/// flow's Logic Apps definition. Only fields we can reliably extract from
/// the generated JSON are populated; per-day work hours are encoded inside
/// nested-if expressions that aren't reverse-parsed in v1.
/// </summary>
public sealed class CloudScheduleDefinitionResult
{
    public CloudScheduleDefinitionResult(
        PowerAutomateOutcome outcome,
        string message,
        string? flowDisplayName,
        IReadOnlyList<DayOfWeek>? workDays,
        int? triggerHour,
        int? triggerMinute,
        string? triggerTimeZone,
        IReadOnlyDictionary<DayOfWeek, CloudDaySchedule>? perDaySchedule = null,
        string? sidecarTimeZone = null,
        string? sidecarGeneratedAt = null,
        DateTimeOffset? triggerStartTimeUtc = null)
    {
        Outcome = outcome;
        Message = message;
        FlowDisplayName = flowDisplayName;
        WorkDays = workDays ?? Array.Empty<DayOfWeek>();
        TriggerHour = triggerHour;
        TriggerMinute = triggerMinute;
        TriggerTimeZone = triggerTimeZone;
        PerDaySchedule = perDaySchedule;
        SidecarTimeZone = sidecarTimeZone;
        SidecarGeneratedAt = sidecarGeneratedAt;
        TriggerStartTimeUtc = triggerStartTimeUtc;
    }

    public PowerAutomateOutcome Outcome { get; }
    public string Message { get; }
    public string? FlowDisplayName { get; }
    public IReadOnlyList<DayOfWeek> WorkDays { get; }
    public int? TriggerHour { get; }
    public int? TriggerMinute { get; }
    public string? TriggerTimeZone { get; }

    /// <summary>
    /// One-shot Recurrence trigger anchor (vacation flows) — the absolute
    /// UTC instant at which Power Automate will fire the trigger once.
    /// Null for weekly-style Cloud Schedule flows (those use hour/minute
    /// + weekDays instead).
    /// </summary>
    public DateTimeOffset? TriggerStartTimeUtc { get; }

    /// <summary>
    /// Per-day work schedule recovered from the sidecar parameter
    /// <c>_oofmgr_source.defaultValue.days</c> we stamp into newer
    /// packages. Null on older flows imported before the sidecar shipped
    /// — those callers see workday + trigger-time compare only.
    /// </summary>
    public IReadOnlyDictionary<DayOfWeek, CloudDaySchedule>? PerDaySchedule { get; }
    public string? SidecarTimeZone { get; }
    public string? SidecarGeneratedAt { get; }
}

public sealed class CloudDaySchedule
{
    public CloudDaySchedule(bool isWorkday, TimeSpan start, TimeSpan end)
    {
        IsWorkday = isWorkday;
        Start = start;
        End = end;
    }
    public bool IsWorkday { get; }
    public TimeSpan Start { get; }
    public TimeSpan End { get; }
}

public interface IPowerAutomateService
{
    Task<PowerAutomateStatusResult> GetOofManagerFlowStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default, IProgress<string>? progress = null);
    Task<PowerAutomateResult> DisableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default);
    Task<PowerAutomateResult> EnableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default);

    /// <summary>
    /// Fetches the deployed cloud flow's Logic Apps definition and decodes
    /// the workdays + trigger time. Powers the "Compare with cloud" feature
    /// in the Weekly mode panel — lets the user see at a glance whether the
    /// local UI matches what's actually running in M365.
    /// </summary>
    Task<CloudScheduleDefinitionResult> GetCloudScheduleDefinitionAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default);

    /// <summary>
    /// Imports the prebuilt OofManager Cloud Schedule solution zip into the
    /// signed-in user's owned Power Platform environment (the one whose
    /// principal owner email matches <paramref name="upnHint"/>), using the
    /// Power Platform CLI (<c>pac solution import</c>) after environment
    /// discovery through the bundled PowerApps PowerShell module. Skips the
    /// manual "download zip → upload via browser" round-trip when pac is
    /// installed and authenticated.
    /// </summary>
    Task<CloudScheduleImportResult> ImportCloudScheduleSolutionAsync(
        string solutionZipPath,
        string solutionUniqueName,
        Guid workflowId,
        string expectedFlowDisplayName,
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        IProgress<string>? progress = null,
        CancellationToken ct = default);
}
