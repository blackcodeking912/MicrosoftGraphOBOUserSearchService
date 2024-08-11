using Azure.Identity;
using Microsoft.Graph.Models;
using MicrosoftGraphOBOUserSearchService.Models;

namespace MicrosoftGraphOBOUserSearchService.Services.Contracts;

public interface IMicrosoftGraphService
{
    void InitializeMicrosoftGraphForOBOUserOAuth(AppSettings settings, Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt);
    Task<User?> GetUserAsync();
    Task<string> GetUserTokenAsync();
    Task SearchMicrosoftGraphAsync(string searchTerm, bool useBetaEndpoint);
}
