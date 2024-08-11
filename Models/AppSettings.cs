using Microsoft.Extensions.Configuration;

namespace MicrosoftGraphOBOUserSearchService.Models
{
    public class AppSettings
    {
        public string? ClientId { get; set; }
        public string? TenantId { get; set; }
        /// <summary>
        /// Example: [ "user.read", "files.read"]
        /// </summary>
        public string[]? GraphUserScopes { get; set; }
        public bool RunDeviceCodeFlow { get; set; }
        /// <summary>
        /// Sting of scopes for TokenRequestContext separated by spaces.
        /// Example: "user.read files.read"
        /// </summary>
        public string? TokenRequestContextScopes { get; set; }

        public static AppSettings LoadConfigurationSettings()
        {
            // Load settings
            var config = new ConfigurationBuilder()
                // appsettings.json is required
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                // appsettings.Development.json" is optional, values override appsettings.json
                .AddJsonFile($"appsettings.Development.json", optional: true, reloadOnChange: true)
                // User secrets are optional, values override both JSON files
                .AddUserSecrets<Program>()
                .Build();

            return config.GetRequiredSection("AppSettings").Get<AppSettings>() ??
                throw new Exception("Could not load app settings. See README for configuration instructions");
        }
    }
}
