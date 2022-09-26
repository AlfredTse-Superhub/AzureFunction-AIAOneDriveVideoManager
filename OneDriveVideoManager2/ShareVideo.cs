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
        private readonly string _functionName = "ShareVideo";

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
                List<ListItem> agentCheckerPairs = await GetAgentCheckerPairs(
                    log,
                    graphClient,
                    functionRunLog);

                List<ErrorLog> errorLogs = new List<ErrorLog>();
                await ManageAgentGroupVideo(
                    log,
                    graphClient,
                    agentCheckerPairs,
                    errorLogs,
                    functionRunLog);

                if (errorLogs.Count > 0)
                {
                    string targetListName = Environment.GetEnvironmentVariable("APPSETTING_ErrorListName2");
                    await LoggingService.CreateErrorLogToSP(
                        log,
                        graphClient,
                        _functionName,
                        targetListName,
                        errorLogs,
                        functionRunLog);
                }
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


        private async Task<List<ListItem>> GetAgentCheckerPairs(
            ILogger log,
            GraphServiceClient graphClient,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                functionRunLog.LastStep = "Get agent & checker pairs from SharePoint";
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
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();

                if (!agentCheckerRequest.Any())
                {
                    functionRunLog.Status = "Succeeded";
                    throw new Exception("No agent-checker pairs are found in SP list");
                }
                // Page through collections
                List<ListItem> agentCheckerPairs = new List<ListItem>();
                agentCheckerPairs.AddRange(agentCheckerRequest.CurrentPage);
                while (agentCheckerRequest.NextPageRequest != null)
                {
                    var nextPage = await agentCheckerRequest.NextPageRequest.GetAsync();
                    agentCheckerPairs.AddRange(nextPage);
                }

                log.LogCritical("SUCCEEDED: get agent & checker pairs.");

                return agentCheckerPairs;

            }
            catch (Exception ex)
            {
                log.LogError($"Function terminated: current step= {functionRunLog.LastStep}");
                if (functionRunLog.Status != "Succeeded")
                {
                    functionRunLog.Status = "Failed";
                }
                functionRunLog.Details = ex.Message;
                throw;
            }
        }


        private async Task ManageAgentGroupVideo(
            ILogger log,
            GraphServiceClient graphClient,
            List<ListItem> agentCheckerPairs,
            List<ErrorLog> errorLogs,
            FunctionRunLog functionRunLog
        )
        {
            functionRunLog.LastStep = "Manage agentGroups' videos";
            await Parallel.ForEachAsync(agentCheckerPairs, async (agentCheckerPair, cancellationToken) =>
            {   
                string pairId = agentCheckerPair.Id ?? "";
                string targetListName = agentCheckerPair.Fields.AdditionalData["SiteTitle"]?.ToString();
                string agentMail = agentCheckerPair.Fields.AdditionalData["AgentMail"]?.ToString();
                string checkerMail = agentCheckerPair.Fields.AdditionalData["CheckerMail"]?.ToString();

                try
                {
                    if (string.IsNullOrWhiteSpace(targetListName) || 
                        string.IsNullOrWhiteSpace(agentMail) || 
                        string.IsNullOrWhiteSpace(checkerMail))
                    {
                        throw new Exception($"Invalid information in AgentAndChecker pair: id={pairId}.");
                    }
                    // Get agentGroup info and members
                    var agentGroup = await graphClient.Groups
                        .Request()
                        .Filter($"mail eq \'{agentMail}\'")
                        .WithMaxRetry(_maxRetry)
                        .GetAsync();
                    if (!agentGroup.CurrentPage.Any())
                    {
                        throw new Exception($"Invalid information for agentGroup('{agentMail}') in AgentAndChecker pair: id={pairId}.");
                    }
                    string agentGroupId = agentGroup?.CurrentPage?.FirstOrDefault()?.Id;
                    var agentGroupMembersRequest = await graphClient.Groups[agentGroupId].Members
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .GetAsync();

                    // Page through collections
                    List<DirectoryObject> agentGroupMembers = new List<DirectoryObject>();
                    agentGroupMembers.AddRange(agentGroupMembersRequest.CurrentPage);
                    while (agentGroupMembersRequest.NextPageRequest != null)
                    {
                        var nextPage = await agentGroupMembersRequest.NextPageRequest.GetAsync();
                        agentGroupMembers.AddRange(nextPage);
                    }

                    bool hasSharingError = false;
                    foreach (User agent in agentGroupMembers) // loop all members
                    {
                        string sharingErrors = "";
                        try
                        {
                            var agentDetails = await graphClient.Users[agent.Id].Drive
                                .Request()
                                .WithMaxRetry(_maxRetry)
                                .GetAsync();
                            string agentDriveId = agentDetails.Id;
                            var agentRecordingsRequest = await graphClient.Drives[agentDriveId].Root.ItemWithPath("/Recordings").Children
                                .Request()
                                .WithMaxRetry(_maxRetry)
                                .GetAsync();

                            // Create 'Shared Recordings' folder if not created before
                            var sharedRequest = await graphClient.Drives[agentDriveId].Root.Children
                                .Request().Filter("name eq \'Shared Recordings\'")
                                .WithMaxRetry(_maxRetry)
                                .GetAsync();
                            if (!sharedRequest.CurrentPage.Any())
                            {
                                var stream = new DriveItem { Name = "Shared Recordings", Folder = new Folder() };
                                await graphClient.Drives[agentDriveId].Root.Children
                                    .Request()
                                    .WithMaxRetry(_maxRetry)
                                    .AddAsync(stream);
                            }

                            // Share videos in 'Shared Recordings' folder
                            var sharedFileRequest = await graphClient.Drives[agentDriveId].Root.ItemWithPath("/Shared Recordings")
                                .Request()
                                .WithMaxRetry(_maxRetry)
                                .GetAsync();
                            sharingErrors = await ShareVideosInOneDrive(
                                log,
                                graphClient,
                                agentRecordingsRequest,
                                targetListName,
                                checkerMail,
                                agentDriveId,
                                sharedFileRequest.Id,
                                agent
                            );
                            if (!string.IsNullOrWhiteSpace(sharingErrors))
                            {
                                throw new Exception("\n One or more video sharing failed:");
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!hasSharingError)
                            {
                                hasSharingError = true;
                            }
                            // Add errorLog if sharing errors occur
                            errorLogs.Add(new ErrorLog
                            {
                                FunctionName = "ShareVideo",
                                StaffName = agent.DisplayName,
                                StaffEmail = agentMail,
                                Details = ex.Message + "\n" + sharingErrors
                            });
                            log.LogError($"One or more onedrive operation failed. Information: Agent name={agent.DisplayName}, id={agent.Id} \n {ex.Message}");
                        }
                    }
                    if (hasSharingError)
                    {
                        functionRunLog.Details += "\n One or more onedrive operation failed, please check SP list 'ShareVideoErrorLog' for reference.\n";
                    }

                }
                catch (Exception ex)
                {
                    if (string.IsNullOrWhiteSpace(functionRunLog.Details))
                    {
                        functionRunLog.Details = "\n One or more agentGroup operations, please check SP list 'ShareVideoErrorLog' for reference.\n";
                    }
                    errorLogs.Add(new ErrorLog
                    {
                        FunctionName = "ShareVideo",
                        StaffName = agentMail,
                        StaffEmail = agentMail,
                        Details = $"Failed to fetch information for agentGroup: '{agentMail}' \n {ex.Message} \n {ex.InnerException?.Message ?? ""}"
                    });
                    log.LogError($"FAILED: manage videos for agentGroup: {agentMail}");
                }
            });
            log.LogCritical("SUCCEEDED: agentGroups' videos' access updated.");
        }


        private async Task<string> ShareVideosInOneDrive(
            ILogger log,
            GraphServiceClient graphClient,
            IDriveItemChildrenCollectionPage agentRecordingsRequest,
            string targetListName,
            string checkerMail,
            string agentDriveId,
            string sharedFileId,
            User agent,
            string errors = ""
        )
        {
            string hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
            string tenantURL = Environment.GetEnvironmentVariable("APPSETTING_TenantURL");
            string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");

            // Page through collections
            List<DriveItem> agentRecordings = new List<DriveItem>();
            agentRecordings.AddRange(agentRecordingsRequest);
            while (agentRecordingsRequest.NextPageRequest != null)
            {
                var nextPage = await agentRecordingsRequest.NextPageRequest.GetAsync();
                agentRecordings.AddRange(nextPage);
            }

            // Share access for each video in an agent's onedrive
            await Parallel.ForEachAsync(agentRecordings, async (video, cancellationToken) => {
                try
                {
                    // Share video: read right
                    List<DriveRecipient> driveRecipient = new List<DriveRecipient>()
                        {
                            new DriveRecipient
                            {
                                Email = checkerMail
                            }
                        };
                    string message = "Here's the file that we're collaborating on.";
                    bool requireSignIn = true;
                    bool sendInvitation = true;
                    var roles = new List<String>()
                        {
                            "read"
                        };
                    await graphClient.Drives[agentDriveId].Items[video.Id]
                        .Invite(driveRecipient, requireSignIn, roles, sendInvitation, message, null)
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .PostAsync();

                    // Create new item in SP list
                    string agentMailSpecial = agent.Mail.ToLower().Replace(".", "_").Replace("@", "_");
                    string itemLink = $"{tenantURL}/personal/{agent.Mail.ToLower().Replace(".", "_").Replace("@", "_")}/Documents/Shared/{video.Name}";
                    TimeSpan t = TimeSpan.FromMilliseconds((double)video.Video.Duration);
                    string formattedDuration = string.Format("{0:D2}:{1:D2}:{2:D2}",
                                    t.Hours,
                                    t.Minutes,
                                    t.Seconds);
                    ListItem newItem = new ListItem
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
                    await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Lists[targetListName].Items
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .AddAsync(newItem);

                    // Update video reference
                    DriveItem videoNewRoot = new DriveItem
                    {
                        ParentReference = new ItemReference
                        {
                            Id = sharedFileId
                        },
                        Name = video.Name,
                    };
                    await graphClient.Drives[agentDriveId].Items[video.Id]
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .UpdateAsync(videoNewRoot);

                }
                catch (Exception ex)
                {
                    errors += $"\n Video name={video.Name}";
                    log.LogError($"FAILED: share access for video '{video.Name}' \n{ex.Message}");
                }
            });

            return errors;
        }
    }
}
