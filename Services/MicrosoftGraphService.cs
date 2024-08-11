using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Search.Query;
using MicrosoftGraphOBOUserSearchService.Models;
using MicrosoftGraphOBOUserSearchService.Services.Contracts;

namespace MicrosoftGraphOBOUserSearchService;

#pragma warning disable CS8602 // Dereference of a possibly null reference.
public class MicrosoftGraphService : IMicrosoftGraphService
{
    private static AppSettings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    private static InteractiveBrowserCredential? _interactiveBrowserCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _graphClient;
    private static Microsoft.Graph.Beta.GraphServiceClient? _betaGraphClient;

    // To initialize your graphClient for OBO User Token request. User
    public void InitializeMicrosoftGraphForOBOUserOAuth(AppSettings settings, Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));

        if (_settings.RunDeviceCodeFlow)
        {
            _deviceCodeCredential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
            {
                ClientId = settings.ClientId,
                TenantId = settings.TenantId,
                DeviceCodeCallback = deviceCodePrompt,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            });
            _graphClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        }
        else
        {
            _interactiveBrowserCredential = new InteractiveBrowserCredential(
                tenantId: settings.TenantId,
                clientId: settings.ClientId,
                options: new InteractiveBrowserCredentialOptions
                {
                    // https://login.microsoftonline.us
                    TenantId = settings.TenantId,
                    ClientId = settings.ClientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    // MUST be http://localhost or http://localhost:PORT
                    // See https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/System-Browser-on-.Net-Core
                    RedirectUri = new Uri(uriString: "http://localhost")
                });

            _graphClient = new GraphServiceClient(
                tokenCredential: _interactiveBrowserCredential,
                scopes: settings.GraphUserScopes,
                baseUrl: "https://graph.microsoft.com/v1.0");
        }
    }

    public async Task<string> GetUserTokenAsync()
    {
        TokenRequestContext tokenRequestContext = new TokenRequestContext(scopes: [$"https://graph.microsoft.com/{_settings.TokenRequestContextScopes}"]);
        AccessToken? response;

        if (_deviceCodeCredential != default)
        {
            response = await _deviceCodeCredential.GetTokenAsync(tokenRequestContext);
            CheckAccessTokenValueNotNull(response);
            return response.Value.Token;
        }
        else
        {
            response = await _interactiveBrowserCredential.GetTokenAsync(requestContext: tokenRequestContext);
            CheckAccessTokenValueNotNull(response);
            return response.Value.Token;
        }
    }

    private void CheckAccessTokenValueNotNull(AccessToken? response)
    {
        if (!response.HasValue)
        {
            throw new InvalidOperationException(message: "Invalid access response");
        }
    }

    public Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _graphClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");

        return _graphClient.Me.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
        });
    }

    public async Task SearchMicrosoftGraphAsync(string searchTerm, bool useBetaEndpoint)
    {
        if (!useBetaEndpoint)
        {
            await CallGraphClientAsync(searchTerm);
        }
        else
        {
            await CallGraphWithBetaClientAsync(searchTerm);
        }
    }

    private async Task CallGraphClientAsync(string searchTerm)
    {
        // Ensure client isn't null
        _ = _graphClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");

        try
        {
            QueryPostRequestBody? requestBody = CreateQueryPostRequestBodyBasedOnSearchTerms<QueryPostRequestBody>(searchTerm, useBetaEndpoint: false);
            QueryPostResponse? queryResponse = await _graphClient?.Search?.Query?.PostAsQueryPostResponseAsync(body: requestBody);
            List<SearchHit> searchHits = [];

            if (queryResponse?.Value != default && queryResponse.Value.Count > 0)
            {
                IEnumerable<List<SearchHitsContainer>?>? hitContainers = queryResponse.Value.Where(predicate: h => h?.HitsContainers?.Count > 0).Select(selector: h => h.HitsContainers);
                foreach (List<SearchHitsContainer>? searchHitsContainers in hitContainers ?? [])
                {
                    foreach (SearchHitsContainer searchHitsContainer in searchHitsContainers ?? Enumerable.Empty<SearchHitsContainer>())
                    {
                        if (searchHitsContainer?.Hits != default && searchHitsContainer.Hits.Count > 0)
                        {
                            searchHits.AddRange(collection: searchHitsContainer.Hits);
                        }
                    }
                }
            }

            DisplaySearchResults(searchHits);
        }
        catch (ODataError ode)
        {
            //Access to ListItem in Graph API requires the following permissions: Sites.Read.All or Sites.ReadWrite.All.
            //However, the application only has the following permissions granted: Mail.Read, User.Read
            throw new Exception(message: ode?.Message ?? ode?.InnerException?.Message);
        }
        catch (Exception ex)
        {
            //Access to ListItem in Graph API requires the following permissions: Sites.Read.All or Sites.ReadWrite.All.
            //However, the application only has the following permissions granted: Mail.Read, User.Read
            throw new Exception(message: ex?.Message ?? ex?.InnerException?.Message);
        }
    }

    private async Task CallGraphWithBetaClientAsync(string searchTerm)
    {
        // Ensure client isn't null
        _ = _graphClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");

        try
        {
            Microsoft.Graph.Beta.Search.Query.QueryPostRequestBody? requestBody = CreateQueryPostRequestBodyBasedOnSearchTerms<Microsoft.Graph.Beta.Search.Query.QueryPostRequestBody>(searchTerm, useBetaEndpoint: false);
            Microsoft.Graph.Beta.Search.Query.QueryPostResponse? queryResponse = await _betaGraphClient?.Search?.Query?.PostAsQueryPostResponseAsync(body: requestBody);
            List<Microsoft.Graph.Beta.Models.SearchHit> searchHits = [];

            if (queryResponse?.Value != default && queryResponse.Value.Count > 0)
            {
                IEnumerable<List<Microsoft.Graph.Beta.Models.SearchHitsContainer>?>? hitContainers = queryResponse.Value.Where(predicate: h => h?.HitsContainers?.Count > 0).Select(selector: h => h.HitsContainers);
                foreach (List<Microsoft.Graph.Beta.Models.SearchHitsContainer>? searchHitsContainers in hitContainers ?? Enumerable.Empty<List<Microsoft.Graph.Beta.Models.SearchHitsContainer>>())
                {
                    foreach (Microsoft.Graph.Beta.Models.SearchHitsContainer searchHitsContainer in searchHitsContainers ?? []) //Enumerable.Empty<Microsoft.Graph.Beta.Models.SearchHitsContainer>()
                    {
                        if (searchHitsContainer?.Hits != default && searchHitsContainer.Hits.Count > 0)
                        {
                            searchHits.AddRange(collection: searchHitsContainer.Hits);
                        }
                    }
                }
            }

            DisplaySearchResults(searchHits);
        }
        catch (Exception ex)
        {
            //Access to ListItem in Graph API requires the following permissions: Sites.Read.All or Sites.ReadWrite.All.
            //However, the application only has the following permissions granted: Mail.Read, User.Read
            throw new Exception(message: ex?.Message ?? ex?.InnerException?.Message);
        }
    }

    private T CreateQueryPostRequestBodyBasedOnSearchTerms<T>(string searchTerm, bool useBetaEndpoint)
    {
        if (!useBetaEndpoint)
        {
            QueryPostRequestBody? requestBody = new QueryPostRequestBody
            {
                Requests = new List<SearchRequest>
                {
                    // Search all content in OneDrive and SharePoint
                    // Queries all content in OneDrive and SharePoint sites to which the signed-in user has read access.
                    // The resource property in the response returns matches that are files and folders as driveItem objects,
                    // matches that are containers (SharePoint lists) as list, and all other matches as listItem.
                    new SearchRequest
                    {
                        EntityTypes = new List<EntityType?>
                        {
                            EntityType.DriveItem,
                            //Access to Site in Graph API requires the following permissions: Sites.Read.All [LP] or Sites.ReadWrite.All [HP].
                            //However, the application only has the following permissions granted: Mail.Read, User.Read
                            //EntityType.Site,
                            EntityType.List,
                            //Access to ListItem in Graph API requires the following permissions: Sites.Read.All [LP] or Sites.ReadWrite.All [HP].
                            //However, the application only has the following permissions granted: Mail.Read, User.Read
                            EntityType.ListItem
                            //EntityType.Drive
                        },
                        Query = new SearchQuery
                        {
                            QueryString = $"{ searchTerm }" ?? string.Empty
                            //QueryTemplate = "{ searchTerm } CreatedBy:Ni'Ko",
                        },
                        From = 0,
                        Size = 500 // max 500
                    },
                },
            };

            return (T)Convert.ChangeType(value: requestBody, conversionType: typeof(T));
        }
        else
        {
            Microsoft.Graph.Beta.Search.Query.QueryPostRequestBody? requestBody = new Microsoft.Graph.Beta.Search.Query.QueryPostRequestBody
            {
                Requests = new List<Microsoft.Graph.Beta.Models.SearchRequest>
                {
                    // Search all content in OneDrive and SharePoint
                    // This example queries all the content in OneDrive and SharePoint sites to which the signed-in user has read access.
                    // The resource property in the response returns matches that are files and folders as driveItem objects,
                    // matches that are containers (SharePoint lists) as list, and all other matches as listItem.
                    new Microsoft.Graph.Beta.Models.SearchRequest
                    {
                        EntityTypes = new List<Microsoft.Graph.Beta.Models.EntityType?>
                        {
                            Microsoft.Graph.Beta.Models.EntityType.List,
                            Microsoft.Graph.Beta.Models.EntityType.ListItem,
                            Microsoft.Graph.Beta.Models.EntityType.Site,
                            Microsoft.Graph.Beta.Models.EntityType.Drive,
                            Microsoft.Graph.Beta.Models.EntityType.DriveItem
                        },
                        Query = new Microsoft.Graph.Beta.Models.SearchQuery
                        {
                            QueryString = $"{ searchTerm }" ?? string.Empty
                        },
                        From = 0,
                        Size = 500, // max 500
                    },
                },
            };
            return (T)Convert.ChangeType(value: requestBody, conversionType: typeof(T));
        }
    }

    private void DisplaySearchResults<T>(List<T> searchHits)
    {
        List<SearchHit>? v1SearchHits = searchHits as List<SearchHit>;
        List<Microsoft.Graph.Beta.Models.SearchHit>? betaSearchHits = searchHits as List<Microsoft.Graph.Beta.Models.SearchHit>;

        if (v1SearchHits != default)
        {
            int index = 0;
            //searchHits.
            v1SearchHits.ForEach(action: (SearchHit searchHit) =>
            {
                DriveItem? driveItemResource = searchHit.Resource as DriveItem;
                Console.WriteLine($"DataSource: {driveItemResource?.OdataType ?? "UNKNOWN SOURCE"}");
                Console.WriteLine($"\tTotalHits: {v1SearchHits.Count}");
                Console.WriteLine($"\tIndex: {index++}");
                Console.WriteLine($"\tRank: {searchHit.Rank}");
                Console.WriteLine($"\tTitle: {driveItemResource?.Name}");
                Console.WriteLine($"\tCreatedByUser: {driveItemResource?.CreatedByUser?.UserPrincipalName ?? driveItemResource?.CreatedByUser?.DisplayName}");
                Console.WriteLine($"\tLastModifiedByUser: {driveItemResource?.LastModifiedByUser?.UserPrincipalName ?? driveItemResource?.LastModifiedByUser?.DisplayName}");
                Console.WriteLine($"\tSummary: {searchHit.Summary}");
                Console.WriteLine($"\tLocation: {driveItemResource?.WebUrl}");
            });
        }

        if (betaSearchHits != default)
        {
            int index = -0;
            //searchHits.
            betaSearchHits.ForEach(action: (Microsoft.Graph.Beta.Models.SearchHit searchHit) =>
            {
                Microsoft.Graph.Beta.Models.DriveItem? driveItemResource = searchHit.Resource as Microsoft.Graph.Beta.Models.DriveItem;
                Console.WriteLine($"DataSource: {driveItemResource?.OdataType ?? "UNKNOWN SOURCE"}");
                Console.WriteLine($"\tTotalHits: {betaSearchHits.Count}");
                Console.WriteLine($"\tIndex: {index++}");
                Console.WriteLine($"\tRank: {searchHit.Rank}");
                Console.WriteLine($"\tTitle: {driveItemResource?.Name}");
                Console.WriteLine($"\tCreatedByUser: {driveItemResource?.CreatedByUser?.UserPrincipalName ?? driveItemResource?.CreatedByUser?.DisplayName}");
                Console.WriteLine($"\tLastModifiedByUser: {driveItemResource?.LastModifiedByUser?.UserPrincipalName ?? driveItemResource?.LastModifiedByUser?.DisplayName}");
                Console.WriteLine($"\tSummary: {searchHit.Summary}");
                Console.WriteLine($"\tLocation: {driveItemResource?.WebUrl}");
            });
        }
    }
}
#pragma warning restore CS8602 // Dereference of a possibly null reference.
