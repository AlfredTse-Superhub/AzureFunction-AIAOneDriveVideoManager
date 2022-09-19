using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace OneDriveVideoManager2
{
    public class Program
    {
        [FunctionName("ShareVideo")]
        public async Task RunAsync([TimerTrigger("0 * 13 * * *", RunOnStartup = true)]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            try
            {
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
                string siteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_siteRelativePath");
                string checkerListName = Environment.GetEnvironmentVariable("APPSETTING_checkerListName");
                string checkingVideoListName = Environment.GetEnvironmentVariable("APPSETTING_checkingVideosListName");

                var client = GraphClientHelper.ConnectToGraphClient();

                var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("expand", "fields")
                };
                var agentCheckerListRequest = await client.Sites.GetByPath(siteRelativePath, hostName).Lists["AgentAndChecker"].Items.Request(queryOptions).GetAsync();
                ForEachAgentCheckerRelation(client, agentCheckerListRequest).Wait();
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
            }
        }

        private static async Task ForEachAgentCheckerRelation(GraphServiceClient client, IListItemsCollectionPage agentCheckersRequest)
        {
            foreach (var agentchecker in agentCheckersRequest.CurrentPage)
            {
                var targetListName = agentchecker.Fields.AdditionalData["SiteTitle"].ToString();
                try
                {
                    ShareItemAccess.Share(client, agentchecker.Fields.AdditionalData["AgentMail"].ToString(), agentchecker.Fields.AdditionalData["CheckerMail"].ToString(), targetListName).Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            if (agentCheckersRequest.NextPageRequest != null)
            {
                ForEachAgentCheckerRelation(client, await agentCheckersRequest.NextPageRequest.GetAsync()).Wait();
            }
        }
    }
}
