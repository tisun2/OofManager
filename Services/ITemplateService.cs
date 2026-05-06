using OofManager.Wpf.Models;

namespace OofManager.Wpf.Services;

public interface ITemplateService
{
    Task<List<OofTemplate>> GetAllTemplatesAsync();
    Task<OofTemplate?> GetTemplateAsync(int id);
    Task SaveTemplateAsync(OofTemplate template);
    Task DeleteTemplateAsync(int id);
}
