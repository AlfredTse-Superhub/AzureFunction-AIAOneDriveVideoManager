using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDriveVideoManager2
{
    public static class ShareItemAccess
    {
        private static string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
        private static string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_spSiteRelativePath");
        public static async Task Share(GraphServiceClient client, string agentMail, string checkerMail, string targetListName, ILogger log)
        {
            try
            {
                var agentGroup = await client.Groups.Request().Filter($"mail eq \'{agentMail}\'").GetAsync();
                var agentGroupId = agentGroup?.CurrentPage?.FirstOrDefault()?.Id;
                var agentGroupMembers = await client.Groups[agentGroupId].Members.Request().GetAsync();
                ForeachMemberInMemberGroup(client, agentGroupMembers, checkerMail, targetListName, log).Wait();
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                // System.Environment.Exit(1);
            }
        }
        private static async Task HandleVideosInOneDrive(
            ILogger log,
            GraphServiceClient client, 
            IDriveItemChildrenCollectionPage saleDriveRecordingsFile, 
            string checkerMail, 
            string agentDriveId, 
            User sale, 
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
                    string itemLink = $"https://m365x71250929-my.sharepoint.com/personal/{sale.Mail.ToLower().Replace(".", "_").Replace("@", "_")}/Documents/Shared/{video.Name}";

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
                    await client.Sites.GetByPath(spSiteRelativePath, hostName).Lists[targetListName].Items.Request().AddAsync(newItem);
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
                HandleVideosInOneDrive(log, client, await saleDriveRecordingsFile.NextPageRequest.GetAsync(), checkerMail, agentDriveId, sale, targetListName, shaedFileId).Wait();
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
    }
}
