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
    public class ShareVideo
    {
        [FunctionName("ShareVideo")]
        public async Task RunAsync([TimerTrigger("0 0 0 * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            try
            {
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_spSiteRelativePath");
                string checkerListName = Environment.GetEnvironmentVariable("APPSETTING_checkerListName");
                string checkingVideoListName = Environment.GetEnvironmentVariable("APPSETTING_checkingVideosListName");

                var client = GraphClientHelper.ConnectToGraphClient();

                var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("expand", "fields")
                };
                var agentCheckerListRequest = await client.Sites.GetByPath(spSiteRelativePath, hostName).Lists["AgentAndChecker"].Items.Request(queryOptions).GetAsync();
                ForEachAgentCheckerRelation(client, agentCheckerListRequest, log).Wait();
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
            }
        }

        private static async Task ForEachAgentCheckerRelation(GraphServiceClient client, IListItemsCollectionPage agentCheckersRequest, ILogger log)
        {
            foreach (var agentchecker in agentCheckersRequest.CurrentPage)
            {
                var targetListName = agentchecker.Fields.AdditionalData["SiteTitle"].ToString();
                try
                {
                    ShareItemAccess.Share(
                        client, 
                        agentchecker.Fields.AdditionalData["AgentMail"].ToString(), 
                        agentchecker.Fields.AdditionalData["CheckerMail"].ToString(), 
                        targetListName,
                        log)
                    .Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            if (agentCheckersRequest.NextPageRequest != null)
            {
                var nextPage = await agentCheckersRequest.NextPageRequest.GetAsync();
                ForEachAgentCheckerRelation(client, nextPage, log).Wait();
            }
        }
    }
}
