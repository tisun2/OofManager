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
        IReadOnlyList<string> flowDisplayNames)
    {
        Outcome = outcome;
        Message = message;
        FlowDisplayNames = flowDisplayNames;
    }

    public PowerAutomateOutcome Outcome { get; }
    public string Message { get; }
    public IReadOnlyList<string> FlowDisplayNames { get; }
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
        IReadOnlyList<string> flowDisplayNames)
    {
        Outcome = outcome;
        State = state;
        Message = message;
        FlowDisplayNames = flowDisplayNames;
    }

    public PowerAutomateOutcome Outcome { get; }
    public PowerAutomateFlowState State { get; }
    public string Message { get; }
    public IReadOnlyList<string> FlowDisplayNames { get; }
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
/// are populated so callers can deep-link the user straight to the imported
/// flow's details page (where they bind the Office 365 connection + flip the
/// flow on — the only two steps the Dataverse Web API can't do silently). On
/// <see cref="CloudScheduleImportOutcome.AlreadyExists"/>, the same env metadata
/// is populated so a confirmation dialog can name the env about to be
/// overwritten before the caller retries with <c>forceOverwrite=true</c>.
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

public interface IPowerAutomateService
{
    Task<PowerAutomateStatusResult> GetOofManagerFlowStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default);
    Task<PowerAutomateResult> DisableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default);
    Task<PowerAutomateResult> EnableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default);

    /// <summary>
    /// Imports the prebuilt OofManager Cloud Schedule solution zip into the
    /// signed-in user's owned Power Platform environment (the one whose
    /// principal owner email matches <paramref name="upnHint"/>) via the
    /// Dataverse Web API. Skips the manual "download zip → upload via
    /// browser" round-trip. Pass <paramref name="forceOverwrite"/>=true to
    /// allow overwriting an existing same-name solution; otherwise the call
    /// returns <see cref="CloudScheduleImportOutcome.AlreadyExists"/> without
    /// importing so the caller can confirm with the user first.
    /// </summary>
    Task<CloudScheduleImportResult> ImportCloudScheduleSolutionAsync(
        string solutionZipPath,
        string solutionUniqueName,
        Guid workflowId,
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        CancellationToken ct = default);
}
