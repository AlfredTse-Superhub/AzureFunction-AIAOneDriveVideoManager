using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;

namespace OneDriveVideoManager.Services
{
    public static class GraphClientHelper
    {
        public static GraphServiceClient ConnectToGraphClient()
        {
            try
            {
                string tenantId = Environment.GetEnvironmentVariable("APPSETTING_TenantId");
                string clientId = Environment.GetEnvironmentVariable("APPSETTING_ClientId");
                string clientSecret = Environment.GetEnvironmentVariable("APPSETTING_ClientSecret");
                string[] scopes = new[] { "https://graph.microsoft.com/.default" };

                TokenCredentialOptions options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
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
