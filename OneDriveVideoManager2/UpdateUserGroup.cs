using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using OneDriveVideoManager2.Models;

namespace OneDriveVideoManager2
{
    public class UpdateUserGroup
    {
        private readonly int _maxRetry = 2;
        private readonly string _functionName = "UpdateUserGroup";

        [FunctionName("UpdateUserGroup")]
        public async Task Run([TimerTrigger("0 0 13 * * *", RunOnStartup = true)]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            FunctionRunLog functionRunLog = new FunctionRunLog { 
                FunctionName = _functionName,
                Details = "",
                Status = "Running",
                LastStep = "Initiate connection"
            };
            GraphServiceClient graphClient = GraphClientHelper.ConnectToGraphClient();

            try
            {
                var excelData = await FetchExcelFromSP(
                    log,
                    graphClient,
                    functionRunLog);

                List<ErrorLog> errorLogs = new List<ErrorLog>();
                await ManageUserGroup(
                    log,
                    graphClient,
                    excelData.userGroups,
                    excelData.aadGroups,
                    errorLogs,
                    functionRunLog);

                if (errorLogs.Count > 0)
                {
                    await CreateErrorLogToSP(
                        log,
                        graphClient,
                        errorLogs,
                        functionRunLog);
                }
            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException.Message ?? ""}");
            }
            finally
            {
                await CreateFunctionRunLogToSP(
                    log,
                    graphClient,
                    functionRunLog);
            }

            log.LogInformation($"C# Timer trigger function ended at: {DateTime.Now}");
        }


        private async Task<(List<UserGroup> userGroups, List<AADGroup> aadGroups)> FetchExcelFromSP(
            ILogger log, 
            GraphServiceClient graphClient, 
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                functionRunLog.LastStep = "Fetch excel from SharePoint";
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_spSiteRelativePath");

                // Fetch excel from SP Documents
                Drive documentLibrary = await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Drive
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();

                var documents = await graphClient.Drives[documentLibrary.Id].Root.Children
                    .Request()
                    .Filter("name eq 'UserGroup.xlsx'")
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();


                // Validation (file existence & modified time)
                if (!documents.Any())
                {
                    functionRunLog.Status = "Succeeded";
                    throw new Exception("No 'UserGroup.xlsx' is found.");
                }
                DriveItem doc = documents.CurrentPage.FirstOrDefault();
                DateTime now = DateTime.Now;
                DateTimeOffset yesterday = new DateTimeOffset(
                    new DateTime(now.Year, now.Month, (now.Day - 1), 0, 0, 0)
                );
                if (doc.LastModifiedDateTime < yesterday)
                {
                    functionRunLog.Status = "Succeeded";
                    throw new Exception("No updates on 'UserGroup.xlsx' is detected on the day before funcion trigger day.");
                }

                // Create List<UserGroup> from excel data
                Stream docStream = await graphClient.Drives[documentLibrary.Id].Items[doc.Id].Content
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();

                log.LogCritical($"SUCCEEDED: Excel file fetched, filename='{doc.Name}'");

                List<UserGroup> userGroups = new List<UserGroup>();
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (docStream)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(docStream))
                    {
                        do
                        {
                            reader.Read(); //Skip header line
                            while (reader.Read()) //Each ROW
                            {
                                UserGroup newUserGroup = new UserGroup();
                                var properties = newUserGroup.GetType().GetProperties();
                                for (int column = 0; column < reader.FieldCount; column++)
                                {
                                    properties[column].SetValue(newUserGroup, reader.GetValue(column), null);
                                }
                                userGroups.Add(newUserGroup);
                            }
                        } while (reader.NextResult()); //Move to NEXT SHEET
                    }
                }

                // Create List<AADGroup> from fetching 'videosharingflow' related groups
                var getGroupsResult = await graphClient.Groups
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();
                List<Group> allGroups = getGroupsResult.CurrentPage.Where(e => e.Description?.ToLower() == "videosharingflow").ToList();
                List<AADGroup> aadGroups = new List<AADGroup>();

                foreach (Group group in allGroups)
                {
                    var memberListResult = await graphClient.Groups[group.Id].Members
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .GetAsync();
                    var memberList = memberListResult.CurrentPage.ToList(); /// paging...
                    aadGroups.Add(new AADGroup()
                    {
                        GroupId = group.Id,
                        GroupName = group.DisplayName,
                        MemberList = memberList
                    });
                }

                log.LogCritical("SUCCEEDED: Excel data processed!");

                return (userGroups, aadGroups);

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

        private async Task ManageUserGroup(
            ILogger log, 
            GraphServiceClient graphClient, 
            List<UserGroup> userGroups, 
            List<AADGroup> aadGroups, 
            List<ErrorLog> errorLogs, 
            FunctionRunLog functionRunLog
        )
        {
            functionRunLog.LastStep = "";
            await Parallel.ForEachAsync(userGroups, async (userGroup, cancellationToken) =>
            {
                try
                {
                    log.LogInformation($"> Processing userGroup: {userGroup.StaffEmail}, {userGroup.AgentGroup}, {userGroup.CheckerGroup}");
                    // validate user
                    var getUserResult = await graphClient.Users
                        .Request()
                        .Filter($"mail eq '{userGroup.StaffEmail}'")
                        .WithMaxRetry(_maxRetry)
                        .GetAsync();
                    var user = getUserResult.CurrentPage.FirstOrDefault() ?? null;
                    if (user == null)
                    {
                        throw new Exception($"User not found, email= '{userGroup.StaffEmail}'");
                    }
                    else
                    {
                        // Update users' group
                        foreach (AADGroup group in aadGroups)
                        {
                            var result = group.MemberList.Where(e => e.Id == user.Id);
                            if (userGroup.AgentGroup == group.GroupName)
                            {
                                if (!result.Any())
                                {
                                    await AddGroupMember(graphClient, user.Id, group.GroupId);
                                }
                            }
                            else if (userGroup.CheckerGroup == group.GroupName)
                            {
                                if (!result.Any())
                                {
                                    await AddGroupMember(graphClient, user.Id, group.GroupId);
                                }
                            }
                            else
                            {
                                if (result.Any())
                                {
                                    await RemoveGroupMember(graphClient, user.Id, group.GroupId);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (string.IsNullOrWhiteSpace(functionRunLog.Details))
                    {
                        functionRunLog.Details = "One or more usergroup update failed, please check SP list 'ErrorLog' for reference";
                    }
                    errorLogs.Add(new ErrorLog
                    {
                        FunctionName = "UpdateUserGroup",
                        StaffName = userGroup.StaffName,
                        StaffEmail = userGroup.StaffEmail,
                        Details = $"{ex.Message} \n{ex.InnerException.Message ?? ""}"
                    });
                }
            });
        }

        private async Task AddGroupMember(GraphServiceClient graphClient, string userId, string groupId)
        {
            var directoryObject = new DirectoryObject
            {
                Id = userId
            };
            await graphClient.Groups[groupId].Members.References
                .Request()
                .WithMaxRetry(_maxRetry)
                .AddAsync(directoryObject);
        }

        private async Task RemoveGroupMember(GraphServiceClient graphClient, string userId, string groupId)
        {
            await graphClient.Groups[groupId].Members[userId].Reference
                .Request()
                .WithMaxRetry(_maxRetry)
                .DeleteAsync();
        }

        private async Task CreateErrorLogToSP(
            ILogger log, 
            GraphServiceClient graphClient, 
            List<ErrorLog> errorLogs,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                functionRunLog.LastStep = "Create ErrorLogs to SP";
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_spSiteRelativePath");
                string targetListName = Environment.GetEnvironmentVariable("APPSETTING_errorListName");

                foreach (ErrorLog errorLog in errorLogs)
                {
                    ListItem newItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                        {
                            {"FunctionName", _functionName},
                            {"StaffName", errorLog.StaffName},
                            {"StaffEmail", errorLog.StaffEmail},
                            {"Details", errorLog.Details}
                        }
                        }
                    };
                    await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Lists[targetListName].Items
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .AddAsync(newItem);
                }
            }
            catch (Exception ex)
            {
                functionRunLog.Details = $"{ex.Message} \n{ex.InnerException.Message ?? ""}";
                log.LogError($"{ex.Message} \n{ex.InnerException.Message ?? ""}");
            }
        }

        private async Task CreateFunctionRunLogToSP(
            ILogger log,
            GraphServiceClient graphClient,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                functionRunLog.LastStep = "Create FunctionRunLog to SP";
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_hostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_spSiteRelativePath");
                string targetListName = Environment.GetEnvironmentVariable("APPSETTING_functionRunLogListName");

                ListItem newItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"FunctionName", functionRunLog.FunctionName},
                            {"Details", functionRunLog.Details},
                            {"Status", functionRunLog.Status},
                            {"LastStep", functionRunLog.LastStep}
                        }
                    }
                };
                await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Lists[targetListName].Items
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .AddAsync(newItem);
            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException.Message ?? ""}");
            }
        }
    }
}
