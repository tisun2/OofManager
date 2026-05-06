using SQLite;

namespace OofManager.Wpf.Models;

public class OofTemplate
{
    [PrimaryKey, AutoIncrement]
    public int Id { get; set; }

    public string Name { get; set; } = string.Empty;
    public string InternalReply { get; set; } = string.Empty;
    public string ExternalReply { get; set; } = string.Empty;
    public bool ExternalAudienceAll { get; set; } = true;
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
