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
}
