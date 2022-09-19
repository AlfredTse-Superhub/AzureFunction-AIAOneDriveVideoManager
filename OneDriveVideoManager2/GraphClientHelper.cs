using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDriveVideoManager2
{
    public static class GraphClientHelper
    {
        public static GraphServiceClient ConnectToGraphClient()
        {
            try
            {
                string tenantId = Environment.GetEnvironmentVariable("APPSETTING_tenantId");
                string clientId = Environment.GetEnvironmentVariable("APPSETTING_clientId");
                string clientSecret = Environment.GetEnvironmentVariable("APPSETTING_clientSecret");
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
                Console.WriteLine(ex);
                System.Environment.Exit(1);
                return null;
            }

        }
    }
}
