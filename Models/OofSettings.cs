namespace OofManager.Wpf.Models;

public enum OofStatus
{
    Disabled,
    Enabled,
    Scheduled
}

public class OofSettings
{
    public OofStatus Status { get; set; } = OofStatus.Disabled;
    public string InternalReply { get; set; } = string.Empty;
    public string ExternalReply { get; set; } = string.Empty;
    public bool ExternalAudienceAll { get; set; } = true;
    public DateTimeOffset? StartTime { get; set; }
    public DateTimeOffset? EndTime { get; set; }
    /// <summary>
    /// Only honoured when <see cref="Status"/> is <see cref="OofStatus.Scheduled"/>.
    /// When true, Exchange will auto-decline new meeting invitations whose
    /// proposed time falls inside the Scheduled OOF window. Used by the
    /// long-vacation feature; the regular work-schedule sync leaves this
    /// false so day-to-day OOF flips don't unexpectedly decline meetings.
    /// </summary>
    public bool DeclineMeetings { get; set; } = false;
}
