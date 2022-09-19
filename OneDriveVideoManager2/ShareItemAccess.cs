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
        private static string siteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_siteRelativePath");
        public static async Task Share(GraphServiceClient client, string agentMail, string checkerMail, string targetListName)
        {
            try
            {
                var getSalesGroup = await client.Groups.Request().Filter($"mail eq \'{agentMail}\'").GetAsync();
                var salesGroupId = getSalesGroup?.CurrentPage?.FirstOrDefault()?.Id;
                var getSalesGroupMembers = await client.Groups[salesGroupId].Members.Request().GetAsync();
                ForEachSaleInSaleGroup(client, getSalesGroupMembers, checkerMail, targetListName).Wait();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                System.Environment.Exit(1);
            }
        }
        private static async Task HandleVideosInOneDrive(GraphServiceClient client, IDriveItemChildrenCollectionPage saleDriveRecordingsFile, string checkerMail, string saleDriveId, User sale, string targetListName, string shaedFileId)
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
                    await client.Drives[saleDriveId].Items[video.Id]
                        .Invite(driveRecipient, requireSignIn, roles, sendInvitation, message, null)
                        .Request()
                        .PostAsync();
                    string itemLink = $"https://m365x71250929-my.sharepoint.com/personal/{sale.Mail.ToLower().Replace(".", "_").Replace("@", "_")}/Documents/Shared/{video.Name}";

                    // Create new item in SP list
                    TimeSpan t = TimeSpan.FromMilliseconds((double)video.Video.Duration);
                    string formattedDuration = string.Format("{0:D2}h:{1:D2}m:{2:D2}s",
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
                    await client.Sites.GetByPath(siteRelativePath, hostName).Lists[targetListName].Items.Request().AddAsync(newItem);
                    var videoNewRoot = new DriveItem
                    {
                        ParentReference = new ItemReference
                        {
                            Id = shaedFileId
                        },
                        Name = video.Name,
                    };
                    await client.Drives[saleDriveId].Items[video.Id].Request().UpdateAsync(videoNewRoot);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fail to sent invitation, {ex}");
                }

            }
            if (saleDriveRecordingsFile.NextPageRequest != null)
            {
                HandleVideosInOneDrive(client, await saleDriveRecordingsFile.NextPageRequest.GetAsync(), checkerMail, saleDriveId, sale, targetListName, shaedFileId).Wait();
            }
        }
        private static async Task ForEachSaleInSaleGroup(GraphServiceClient client, IGroupMembersCollectionWithReferencesPage getSalesGroupMembers, string checkerMail, string targetListName)
        {
            foreach (User sale in getSalesGroupMembers.CurrentPage)
            {
                try
                {
                    var getSaleDetail = await client.Users[sale.Id].Drive.Request().GetAsync();
                    var saleDriveId = getSaleDetail.Id;
                    IDriveItemChildrenCollectionPage saleDriveRecordingsFile = await client.Drives[saleDriveId].Root.ItemWithPath("/Recordings").Children.Request().GetAsync();

                    // Create Shared folder if not created before
                    var sharedRequest = await client.Drives[saleDriveId].Root.Children.Request().Filter("name eq \'Shared\'").GetAsync();
                    if (sharedRequest.CurrentPage.Count == 0)
                    {
                        var stream = new DriveItem { Name = "Shared", Folder = new Folder() };
                        await client.Drives[saleDriveId].Root.Children
                            .Request()
                            .AddAsync(stream);
                    }

                    var sharedFileRequest = await client.Drives[saleDriveId].Root.ItemWithPath("/Shared").Request().GetAsync();
                    var sharedFileId = sharedFileRequest.Id;
                    HandleVideosInOneDrive(client, saleDriveRecordingsFile, checkerMail, saleDriveId, sale, targetListName, sharedFileId).Wait();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{sale.Id} faile to get drive, {ex}");
                }
                Console.WriteLine("");
            }
            if (getSalesGroupMembers.NextPageRequest != null)
            {
                ForEachSaleInSaleGroup(client, await getSalesGroupMembers.NextPageRequest.GetAsync(), checkerMail, targetListName).Wait();
            }
        }


        //    public static async Task Share(GraphServiceClient client, string agentMail, string checkerMail, string checkingVideosListId)
        //    {
        //        try
        //        {
        //            string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
        //            string siteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_siteRelativePath");

        //            var getSalesGroup = await client.Groups.Request().Filter($"mail eq \'{agentMail}\'").GetAsync();
        //            var salesGroupId = getSalesGroup?.CurrentPage?.FirstOrDefault()?.Id;
        //            var getSalesGroupMembers = await client.Groups[salesGroupId].Members.Request().GetAsync();
        //            while (true)
        //            {
        //                foreach (User sale in getSalesGroupMembers.CurrentPage)
        //                {
        //                    try
        //                    {
        //                        var getSaleDetail = await client.Users[sale.Id].Drive.Request().GetAsync();
        //                        var saleDriveId = getSaleDetail.Id;
        //                        var saleDriveRecordingsFile = await client.Drives[saleDriveId].Root.ItemWithPath("/Recordings").Children.Request().GetAsync();

        //                        var sharedRequest = await client.Drives[saleDriveId].Root.Children.Request().Filter("name eq \'Shared\'").GetAsync();
        //                        if (sharedRequest.CurrentPage.Count == 0)
        //                        {
        //                            var stream = new DriveItem { Name = "Shared", Folder = new Folder() };
        //                            await client.Drives[saleDriveId].Root.Children
        //                            .Request()
        //                            .AddAsync(stream);
        //                        }

        //                        var sharedFileRequest = await client.Drives[saleDriveId].Root.ItemWithPath("/Shared").Request().GetAsync();
        //                        var shaedFileId = sharedFileRequest.Id;
        //                        while (true)
        //                        {
        //                            foreach (DriveItem video in saleDriveRecordingsFile)
        //                            {
        //                                Console.WriteLine(video.Name);
        //                                try
        //                                {
        //                                    List<DriveRecipient> driveRecipient = new List<DriveRecipient>()
        //                                    {
        //                                        new DriveRecipient
        //                                        {
        //                                            Email = checkerMail
        //                                        }
        //                                    };
        //                                    var message = "Here's the file that we're collaborating on.";
        //                                    var requireSignIn = true;
        //                                    var sendInvitation = true;
        //                                    var roles = new List<String>()
        //                                    {
        //                                        "read"
        //                                    };
        //                                    await client.Drives[saleDriveId].Items[video.Id]
        //                                        .Invite(driveRecipient, requireSignIn, roles, sendInvitation, message, null)
        //                                        .Request()
        //                                        .PostAsync();
        //                                    string itemLink = $"https://m365x71250929-my.sharepoint.com/personal/{sale.Mail.ToLower().Replace(".", "_").Replace("@", "_")}/Documents/Shared/{video.Name}";

        //                                    var newItem = new ListItem
        //                                    {
        //                                        Fields = new FieldValueSet
        //                                        {
        //                                            AdditionalData = new Dictionary<string, object>()
        //                                            {
        //                                                {"Title", "New video"},
        //                                                {"Checked", false},
        //                                                {"LinkToVideo", itemLink},
        //                                                {"CheckerGroup", checkerMail}
        //                                            }
        //                                        }
        //                                    };
        //                                    await client.Sites.GetByPath(siteRelativePath, hostName).Lists[checkingVideosListId].Items.Request().AddAsync(newItem);
        //                                    var videoNewRoot = new DriveItem
        //                                    {
        //                                        ParentReference = new ItemReference
        //                                        {
        //                                            Id = shaedFileId
        //                                        },
        //                                        Name = video.Name,
        //                                    };
        //                                    await client.Drives[saleDriveId].Items[video.Id].Request().UpdateAsync(videoNewRoot);
        //                                }
        //                                catch (Exception ex)
        //                                {
        //                                    Console.WriteLine($"Fail to sent invitation, {ex}");
        //                                }

        //                            }
        //                            if (saleDriveRecordingsFile.NextPageRequest == null)
        //                            {
        //                                break;
        //                            }
        //                            else
        //                            {
        //                                try
        //                                {
        //                                    saleDriveRecordingsFile = await saleDriveRecordingsFile.NextPageRequest.GetAsync();
        //                                }
        //                                catch (Exception ex)
        //                                {
        //                                    Console.WriteLine($"Fail to get next item in {sale} drive, {ex}");
        //                                }
        //                            }
        //                        }
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine($"{sale.Id} faile to get drive, {ex}");
        //                    }
        //                    Console.WriteLine("");
        //                }
        //                if (getSalesGroupMembers.NextPageRequest == null)
        //                {
        //                    break;
        //                }
        //                else
        //                {
        //                    try
        //                    {
        //                        getSalesGroupMembers = await getSalesGroupMembers.NextPageRequest.GetAsync();
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine($"Faile to get next page of sales, {ex}");
        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex);
        //            System.Environment.Exit(1);
        //        }
        //    }
    }
}
