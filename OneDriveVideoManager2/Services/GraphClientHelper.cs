using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;

namespace OneDriveVideoManager.Services
{
    public static class GraphClientHelper
    {
        public static GraphServiceClient ConnectToGraphClient(bool useTestTenant = false)
        {
            try
            {
                string environment = (useTestTenant) ? "-Test" : "";
                string tenantId = Environment.GetEnvironmentVariable($"APPSETTING_TenantId{environment}");
                string clientId = Environment.GetEnvironmentVariable($"APPSETTING_ClientId{environment}");
                string clientSecret = Environment.GetEnvironmentVariable($"APPSETTING_ClientSecret{environment}");
                string[] scopes = new[] { "https://graph.microsoft.com/.default" };

                TokenCredentialOptions options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                ClientSecretCredential clientSecretCredential = new ClientSecretCredential(
                    tenantId,
                    clientId,
                    clientSecret,
                    options
                );
                return new GraphServiceClient(clientSecretCredential, scopes);
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
