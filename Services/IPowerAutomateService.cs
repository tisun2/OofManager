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
    Task<PowerAutomateStatusResult> GetOofManagerFlowStatusAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, CancellationToken ct = default, IProgress<string>? progress = null);
    Task<PowerAutomateResult> DisableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default);
    Task<PowerAutomateResult> EnableOofManagerFlowsAsync(string? upnHint, string? displayNameHint, string expectedFlowDisplayName, IProgress<string>? progress = null, CancellationToken ct = default);

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
        string? upnHint,
        string? displayNameHint,
        bool forceOverwrite,
        IProgress<string>? progress = null,
        CancellationToken ct = default);
}
