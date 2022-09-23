using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using OneDriveVideoManager.Models;
using OneDriveVideoManager.Services;

namespace OneDriveVideoManager
{
    public class ShareVideo
    {
        private readonly int _maxRetry = 2;
        private readonly string _functionName = "UpdateUserGroup";

        [FunctionName("ShareVideo")]
        public async Task RunAsync([TimerTrigger("0 0 22 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function({_functionName}) executed at: {DateTime.Now}");

            FunctionRunLog functionRunLog = new FunctionRunLog
            {
                FunctionName = _functionName,
                Details = "",
                Status = "Running",
                LastStep = "Initiate connection"
            };
            GraphServiceClient graphClient = GraphClientHelper.ConnectToGraphClient();

            try
            {
                await ManageOnedriveVideos(
                    log,
                    graphClient,
                    functionRunLog
                );
                //string hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
                //string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");

                //var queryOptions = new List<QueryOption>()
                //{
                //    new QueryOption("expand", "fields")
                //};
                //var agentCheckerListRequest = await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Lists["AgentAndChecker"].Items
                //    .Request(queryOptions)
                //    .GetAsync();
                //ForEachAgentCheckerRelation(graphClient, agentCheckerListRequest, log).Wait();
            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
                await MailService.SendReportErrorEmail(
                    log,
                    graphClient,
                    _functionName,
                    functionRunLog.Details);
            }
            finally
            {
                await LoggingService.CreateFunctionRunLogToSP(
                    log,
                    graphClient,
                    functionRunLog);
            }

            log.LogInformation($"C# Timer trigger function({_functionName}) ended at: {DateTime.Now}");
        }


        private async Task<List<BaseItem>> getAllItemsInCollection<BaseItem>(ICollectionPage collectionPage)
        {
            List<BaseItem> itemList = new List<BaseItem>();
            itemList.AddRange(collectionPage.CurrentPage);
            while (collectionPage.NextPageRequest != null)
            {
                var nextPage = await collectionPage.NextPageRequest.GetAsync();
                itemList.AddRange(nextPage);
            }

            return itemList;
        }

        private async Task<List<BaseItem>> getAllItemsInCollection2<T>(
            ICollectionPage<List<BaseItem>> collectionPage,
            IGroupMembersCollectionWithReferencesPage nextPageRequest
        )
        {
            List<BaseItem> itemList = new List<BaseItem>();
            itemList.AddRange(collectionPage.CurrentPage);
            while (collectionPage.IGroupMembersCollectionWithReferencesPage != null)
            {
                var nextPage = await collectionPage.NextPageRequest.GetAsync();
                itemList.AddRange(nextPage);
            }

            return itemList;
        }

        private async Task ManageOnedriveVideos(
            ILogger log,
            GraphServiceClient graphClient,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                functionRunLog.LastStep = "Fetch excel from SharePoint";
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");
                string agentCheckerListName = Environment.GetEnvironmentVariable("APPSETTING_AgentCheckerListName");

                // Get all agent-checker pairs in SP list
                var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("expand", "fields")
                };
                var agentCheckerRequest = await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Lists[agentCheckerListName].Items
                    .Request(queryOptions)
                    .GetAsync();

                if (agentCheckerRequest.Count == 0)
                {
                    functionRunLog.Status = "Succeeded";
                    throw new Exception("No agent-checker pairs are found in SP list");
                }
                // Page through collections
                List<ListItem> agentCheckerList = new List<ListItem>();
                agentCheckerList.AddRange(agentCheckerRequest.CurrentPage);
                while (agentCheckerRequest.NextPageRequest != null)
                {
                    var nextPage = await agentCheckerRequest.NextPageRequest.GetAsync();
                    agentCheckerList.AddRange(nextPage);
                }

                List<ErrorLog> errorLogs = new List<ErrorLog>();
                foreach (ListItem agentCheckerPair in agentCheckerList)
                {
                    string targetListName = agentCheckerPair.Fields.AdditionalData["SiteTitle"].ToString();
                    await ManageAgentGroupVideo(
                        log,
                        graphClient,
                        agentCheckerPair.Fields.AdditionalData["AgentMail"].ToString(),
                        agentCheckerPair.Fields.AdditionalData["CheckerMail"].ToString(),
                        targetListName,
                        errorLogs);
                }
                //if (agentCheckerRequest.NextPageRequest != null)
                //{
                //    var nextPage = await agentCheckerRequest.NextPageRequest.GetAsync();
                //    ForEachAgentCheckerRelation(client, nextPage, log).Wait();
                //}
            }
            catch (Exception ex)
            {
                log.LogError($"Function terminated: current step= {functionRunLog.LastStep}");
                throw;
            }
        }


        private async Task ManageAgentGroupVideo(
            ILogger log,
            GraphServiceClient client,
            string agentMail,
            string checkerMail,
            string targetListName,
            List<ErrorLog> errorLogs
        )
        {
            try
            {
                var agentGroup = await client.Groups.Request().Filter($"mail eq \'{agentMail}\'").GetAsync();
                string agentGroupId = agentGroup?.CurrentPage?.FirstOrDefault()?.Id;
                if (string.IsNullOrWhiteSpace(agentGroupId))
                {
                    // ...?
                }
                var agentGroupMembers = await client.Groups[agentGroupId].Members.Request().GetAsync();
                // ForeachMemberInMemberGroup(client, agentGroupMembers, checkerMail, targetListName, log).Wait();

                foreach (User agent in agentGroupMembers.CurrentPage)
                {
                    try
                    {
                        var agentDetails = await client.Users[agent.Id].Drive.Request().GetAsync();
                        var agentDriveId = agentDetails.Id;
                        IDriveItemChildrenCollectionPage saleDriveRecordingsFile = await client.Drives[agentDriveId].Root.ItemWithPath("/Recordings").Children.Request().GetAsync();

                        // Create Shared folder if not created before
                        var sharedRequest = await client.Drives[agentDriveId].Root.Children.Request().Filter("name eq \'Shared\'").GetAsync();
                        if (sharedRequest.CurrentPage.Count == 0)
                        {
                            var stream = new DriveItem { Name = "Shared", Folder = new Folder() };
                            await client.Drives[agentDriveId].Root.Children
                                .Request()
                                .AddAsync(stream);
                        }

                        var sharedFileRequest = await client.Drives[agentDriveId].Root.ItemWithPath("/Shared").Request().GetAsync();
                        var sharedFileId = sharedFileRequest.Id;
                        HandleVideosInOneDrive(log, client, saleDriveRecordingsFile, checkerMail, agentDriveId, agent, targetListName, sharedFileId).Wait();
                    }
                    catch (Exception ex)
                    {
                        log.LogError($"agent name: {agent.DisplayName}, id: {agent.Id} failed to get onedrive. \n {ex.Message}");
                    }
                }
                if (agentGroupMembers.NextPageRequest != null)
                {
                    ForeachMemberInMemberGroup(client, await agentGroupMembers.NextPageRequest.GetAsync(), checkerMail, targetListName, log).Wait();
                }
            }
            catch (Exception ex)
            {
                // log.LogError(ex.Message);
                log.LogError($"FAILED: update group for user: {userGroup.StaffEmail}");
            }
        }


        private static async Task HandleVideosInOneDrive(
            ILogger log,
            GraphServiceClient client,
            IDriveItemChildrenCollectionPage saleDriveRecordingsFile,
            string checkerMail,
            string agentDriveId,
            User agent,
            string targetListName,
            string shaedFileId
        )
        {
            foreach (Microsoft.Graph.DriveItem video in saleDriveRecordingsFile)
            {
                var hi = video.Video;
                Console.WriteLine(video.Name);
                try
                {
                    // Share access
                    List<DriveRecipient> driveRecipient = new List<DriveRecipient>()
                        {
                            new DriveRecipient
                            {
                                Email = checkerMail
                            }
                        };
                    var message = "Here's the file that we're collaborating on.";
                    var requireSignIn = true;
                    var sendInvitation = true;
                    var roles = new List<String>()
                                        {
                                            "read"
                                        };
                    await client.Drives[agentDriveId].Items[video.Id]
                        .Invite(driveRecipient, requireSignIn, roles, sendInvitation, message, null)
                        .Request()
                        .PostAsync();

                    string agentMailSpecial = agent.Mail.ToLower().Replace(".", "_").Replace("@", "_");
                    string itemLink = $"{_tenantURL}/personal/{agent.Mail.ToLower().Replace(".", "_").Replace("@", "_")}/Documents/Shared/{video.Name}";
                    // Create new item in SP list
                    TimeSpan t = TimeSpan.FromMilliseconds((double)video.Video.Duration);
                    string formattedDuration = string.Format("{0:D2}:{1:D2}:{2:D2}",
                                    t.Hours,
                                    t.Minutes,
                                    t.Seconds);

                    var newItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                                {
                                    {"Title", "New video"},
                                    {"Checked", false},
                                    {"Duration", formattedDuration},
                                    {"LinkToVideo", itemLink}
                                }
                        }
                    };
                    await client.Sites.GetByPath(_spSiteRelativePath, _hostName).Lists[targetListName].Items.Request().AddAsync(newItem);
                    var videoNewRoot = new DriveItem
                    {
                        ParentReference = new ItemReference
                        {
                            Id = shaedFileId
                        },
                        Name = video.Name,
                    };
                    await client.Drives[agentDriveId].Items[video.Id].Request().UpdateAsync(videoNewRoot);
                }
                catch (Exception ex)
                {
                    log.LogError($"Fail to sent invitation, {ex.Message}");
                }
            }
            if (saleDriveRecordingsFile.NextPageRequest != null)
            {
                HandleVideosInOneDrive(log, client, await saleDriveRecordingsFile.NextPageRequest.GetAsync(), checkerMail, agentDriveId, agent, targetListName, shaedFileId).Wait();
            }
        }
        private static async Task ForeachMemberInMemberGroup(
            GraphServiceClient client,
            IGroupMembersCollectionWithReferencesPage agentGroupMembers,
            string checkerMail,
            string targetListName,
            ILogger log
        )
        {
            foreach (User agent in agentGroupMembers.CurrentPage)
            {
                try
                {
                    var agentDetails = await client.Users[agent.Id].Drive.Request().GetAsync();
                    var agentDriveId = agentDetails.Id;
                    IDriveItemChildrenCollectionPage saleDriveRecordingsFile = await client.Drives[agentDriveId].Root.ItemWithPath("/Recordings").Children.Request().GetAsync();

                    // Create Shared folder if not created before
                    var sharedRequest = await client.Drives[agentDriveId].Root.Children.Request().Filter("name eq \'Shared\'").GetAsync();
                    if (sharedRequest.CurrentPage.Count == 0)
                    {
                        var stream = new DriveItem { Name = "Shared", Folder = new Folder() };
                        await client.Drives[agentDriveId].Root.Children
                            .Request()
                            .AddAsync(stream);
                    }

                    var sharedFileRequest = await client.Drives[agentDriveId].Root.ItemWithPath("/Shared").Request().GetAsync();
                    var sharedFileId = sharedFileRequest.Id;
                    HandleVideosInOneDrive(log, client, saleDriveRecordingsFile, checkerMail, agentDriveId, agent, targetListName, sharedFileId).Wait();
                }
                catch (Exception ex)
                {
                    log.LogError($"agent name: {agent.DisplayName}, id: {agent.Id} failed to get onedrive. \n {ex.Message}");
                }
            }
            if (agentGroupMembers.NextPageRequest != null)
            {
                ForeachMemberInMemberGroup(client, await agentGroupMembers.NextPageRequest.GetAsync(), checkerMail, targetListName, log).Wait();
            }
        }


        //private static async Task ForEachAgentCheckerRelation(GraphServiceClient client, IListItemsCollectionPage agentCheckersRequest, ILogger log)
        //{
        //    foreach (var agentchecker in agentCheckersRequest.CurrentPage)
        //    {
        //        var targetListName = agentchecker.Fields.AdditionalData["SiteTitle"].ToString();
        //        try
        //        {
        //            await ShareItemAccess.Share(
        //                client,
        //                agentchecker.Fields.AdditionalData["AgentMail"].ToString(),
        //                agentchecker.Fields.AdditionalData["CheckerMail"].ToString(),
        //                targetListName,
        //                log);
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex);
        //        }
        //    }
        //    if (agentCheckersRequest.NextPageRequest != null)
        //    {
        //        var nextPage = await agentCheckersRequest.NextPageRequest.GetAsync();
        //        ForEachAgentCheckerRelation(client, nextPage, log).Wait();
        //    }
        //}
    }
}
