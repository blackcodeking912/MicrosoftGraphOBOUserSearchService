using Microsoft.Graph.Models;
using MicrosoftGraphOBOUserSearchService;
using MicrosoftGraphOBOUserSearchService.Models;
using MicrosoftGraphOBOUserSearchService.Services.Contracts;

/// <summary>
/// 
/// Microsoft Graph OBO User Search Service [OBO User Delegate Permissions Microsoft Graph]
/// .NET 8.0 Console Application
/// 
/// </summary>
#pragma warning disable CS8602 // Dereference of a possibly null reference.
Console.WriteLine("Start .NET Microsoft Graph OBO User Search Service\n");

AppSettings _settings = AppSettings.LoadConfigurationSettings();
IMicrosoftGraphService _microsoftGraphService = new MicrosoftGraphService();

// Initialize Graph
InitializeGraph(_settings);

// Greet user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. Search graph");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit program
            Console.WriteLine("Goodbye..");
            break;
        case 1:
            // Display user access token [OBO Delegate Microsoft Graph Permissions]
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // Search graph
            await SearchGraphAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again..");
            break;
    }
}

/// <summary>
/// 
/// This method initializes the Microsoft Graph Client for making RestAPI calls to Azure tenant
/// 
/// </summary>
void InitializeGraph(AppSettings settings)
{
    _microsoftGraphService.InitializeMicrosoftGraphForOBOUserOAuth(settings,
        (deviceCodeInfo, cancel) =>
        {
            // Display the device code message to user.
            // This tells them where to go to sign in
            // and provides the code to use.
            Console.WriteLine(deviceCodeInfo.Message);
            return Task.FromResult(0);
        });
}

/// <summary>
/// 
/// Once signed in, this method is used to retrieve the users information from the graph
/// /me endpoint
/// 
/// </summary>
async Task? GreetUserAsync()
{
    try
    {
        User? user = await _microsoftGraphService.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? string.Empty}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

/// <summary>
/// 
/// DisplayAccessTokenAsync
/// 
/// </summary>
async Task DisplayAccessTokenAsync()
{
    try
    {
        var user_access_token = await _microsoftGraphService.GetUserTokenAsync();
        Console.WriteLine($"User access token for downstream api [Microsoft Graph] requests: {user_access_token}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}

/// <summary>
/// 
/// SearchGraphAsync
/// 
/// </summary>
async Task SearchGraphAsync()
{
    Console.WriteLine("enter search term or phrase");
    var searchTerm = Console.ReadLine() ?? throw new Exception("No search term or phrase provided");
    await _microsoftGraphService.SearchMicrosoftGraphAsync(searchTerm, useBetaEndpoint: false);
}
#pragma warning restore CS8602 // Dereference of a possibly null reference.