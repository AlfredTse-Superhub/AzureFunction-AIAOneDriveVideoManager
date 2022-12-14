using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;
using ExcelDataReader.Log;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using OneDriveVideoManager.Models;
using OneDriveVideoManager.Services;

namespace OneDriveVideoManager
{
    public class UpdateUserGroup
    {
        private readonly string _functionName = "UpdateUserGroup";
        private readonly int _maxRetry = 2;
        private Drive _documentLibrary;
        private byte[] _docContent;

        [FunctionName("UpdateUserGroup")]
        public async Task Run([TimerTrigger("%APPSETTING_ScheduleTriggerTime-UpdateGroup%")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function({_functionName}) executed at: {DateTime.Now}");

            FunctionRunLog functionRunLog = new FunctionRunLog { 
                Id = "",
                FunctionName = _functionName,
                Details = "",
                Status = "Running",
                LastStep = "Initiate connection to tenant"
            };

            GraphServiceClient graphClient = GraphClientHelper.ConnectToGraphClient();

            // In case CDX tenant can't send email, use another tenant to test mailing function
            bool sendEmailWithTestTenant = bool.Parse(Environment.GetEnvironmentVariable("APPSETTING_SendEmailWithTestTenant"));
            GraphServiceClient graphClientForMail = (sendEmailWithTestTenant)
                ? GraphClientHelper.ConnectToGraphClient(useTestTenant: true)
                : graphClient;

            try
            {
                await LoggingService.PostFunctionRunLogToSP(
                    log,
                    graphClient,
                    functionRunLog);

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
                    string targetListName = Environment.GetEnvironmentVariable("APPSETTING_ErrorListName1");
                    await LoggingService.CreateErrorLogToSP(
                        log,
                        graphClient,
                        _functionName,
                        targetListName,
                        errorLogs,
                        functionRunLog);
                }
                functionRunLog.UpdatedRecords = functionRunLog.TotalRecords - errorLogs.Count;

                await SaveExcelToSP(
                    log,
                    graphClient,
                    functionRunLog);

            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
                await MailService.SendReportErrorEmail(
                    log,
                    graphClientForMail,
                    _functionName,
                    functionRunLog.Details);
            }
            finally
            {
                await LoggingService.PostFunctionRunLogToSP(
                    log,
                    graphClient,
                    functionRunLog);
            }

            log.LogInformation($"C# Timer trigger function({_functionName}) ended at: {DateTime.Now}");
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
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");

                // Fetch excel from SP Documents
                _documentLibrary = await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Drive
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();

                var documents = await graphClient.Drives[_documentLibrary.Id].Root.Children
                    .Request()
                    .Filter("name eq 'UserGroup.xlsx'")
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();


                // Validation (file existence & modified time)
                if (!documents.Any())
                {
                    functionRunLog.Status = "Succeeded";
                    throw new Exception("'UserGroup.xlsx' is not found.");
                }
                DriveItem doc = documents.CurrentPage.FirstOrDefault();
                DateTime uctTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, (DateTime.Now.Day - 1), 0, 0, 0);
                DateTimeOffset utcYesterday = DateTime.SpecifyKind(uctTime, DateTimeKind.Utc); //yesterday 12:00am
                if (doc.LastModifiedDateTime < utcYesterday)
                {
                    functionRunLog.Status = "Succeeded";
                    throw new Exception("No updates on 'UserGroup.xlsx' is detected on the day before funcion trigger day.");
                }


                // Get file content and save byte[]
                Stream docStream = await graphClient.Drives[_documentLibrary.Id].Items[doc.Id].Content
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();

                MemoryStream ms = new MemoryStream();
                docStream.CopyTo(ms);
                _docContent = ms.ToArray();

                log.LogCritical($"SUCCEEDED: Fetch Excel file '{doc.Name}'.   Ended at: {DateTime.Now}");


                // Process excel data and create new data models
                var excelData = await ProcessExcelData(
                    log,
                    graphClient,
                    docStream,
                    functionRunLog);

                return excelData;

            }
            catch (Exception ex)
            {
                log.LogError($"Function terminated: current step='{functionRunLog.LastStep}'");
                if (functionRunLog.Status != "Succeeded")
                {
                    functionRunLog.Status = "Failed";
                }
                functionRunLog.Details = ex.Message;
                throw;
            }
        }


        private async Task<(List<UserGroup> userGroups, List<AADGroup> aadGroups)> ProcessExcelData(
            ILogger log,
            GraphServiceClient graphClient,
            Stream docStream,
            FunctionRunLog functionRunLog
        )
        {
            functionRunLog.LastStep = "Process excel data.";

            // Create List<UserGroup> from excel data
            List<UserGroup> userGroups = new List<UserGroup>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            int rowNo = 1;
            using (docStream)
            {
                using (var reader = ExcelReaderFactory.CreateReader(docStream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        if (rowNo > 1) //Skip header line
                        {
                            UserGroup newUserGroup = new UserGroup();
                            var properties = newUserGroup.GetType().GetProperties();
                            properties[0].SetValue(newUserGroup, rowNo, null);
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                properties[column + 1].SetValue(newUserGroup, reader.GetValue(column), null);
                            }
                            bool isDuplicated = userGroups.Where(e => e.StaffEmail == newUserGroup.StaffEmail).Any();
                            if (!isDuplicated) //Drop duplicated email row
                            {
                                userGroups.Add(newUserGroup);
                            }
                        }
                        rowNo++;
                    }
                }
            }
            functionRunLog.TotalRecords = rowNo - 1;


            // Create List<AADGroup> from fetching 'videosharingflow' related groups
            var getGroupsResult = await graphClient.Groups
                .Request()
                .WithMaxRetry(_maxRetry)
                .GetAsync();

            List<Group> allGroups = getGroupsResult.CurrentPage.ToList();
            while (getGroupsResult.NextPageRequest != null) // Page through collections
            {
                allGroups.AddRange(await getGroupsResult.NextPageRequest.GetAsync());
            }
            allGroups = allGroups.Where(e => e.Description?.ToLower() == "videosharingflow").ToList();

            List<AADGroup> aadGroups = new List<AADGroup>();
            foreach (Group group in allGroups)
            {
                var memberListResult = await graphClient.Groups[group.Id].Members
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .GetAsync();

                List<DirectoryObject> memberList = memberListResult.CurrentPage.ToList();
                while (memberListResult.NextPageRequest != null) // Page through collections
                {
                    memberList.AddRange(await memberListResult.NextPageRequest.GetAsync());
                }
                aadGroups.Add(new AADGroup()
                {
                    GroupId = group.Id,
                    GroupName = group.DisplayName,
                    MemberList = memberList
                });
            }

            log.LogCritical($"SUCCEEDED: Process excel data.   Ended at: {DateTime.Now}");

            return (userGroups, aadGroups);
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
            functionRunLog.LastStep = "Manage user groups";
            await Parallel.ForEachAsync(userGroups, async (userGroup, cancellationToken) =>
            {
                try
                {
                    log.LogInformation($"> Processing userGroup: {userGroup.StaffEmail}, {userGroup.AgentGroup}, {userGroup.CheckerGroup}");

                    // Validate excel data
                    var getUserResult = await graphClient.Users
                        .Request()
                        .Filter($"mail eq '{userGroup.StaffEmail}'")
                        .WithMaxRetry(_maxRetry)
                        .GetAsync();
                    var user = getUserResult.CurrentPage.FirstOrDefault() ?? null;
                    if (user == null) // validate user
                    {
                        throw new Exception($"User not found.");
                    }
                    bool isAgentGroupValid = aadGroups.Where(e => e.GroupName == userGroup.AgentGroup).Any();
                    bool isCheckerGroupValid = aadGroups.Where(e => e.GroupName == userGroup.CheckerGroup).Any();
                    if ((!string.IsNullOrWhiteSpace(userGroup.AgentGroup) && !isAgentGroupValid) 
                        || (!string.IsNullOrWhiteSpace(userGroup.CheckerGroup) && !isCheckerGroupValid)) // validate userGroup name
                    {
                        throw new Exception($"Invalid usergroup name ({userGroup.AgentGroup}/{userGroup.CheckerGroup}).");
                    }

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
                catch (Exception ex)
                {
                    if (string.IsNullOrWhiteSpace(functionRunLog.Details))
                    {
                        functionRunLog.Details = "One or more usergroup update failed, please check SP list 'UpdateUserGroupErrorLog' for reference.\n";
                    }
                    errorLogs.Add(new ErrorLog
                    {
                        FunctionName = "UpdateUserGroup",
                        StaffName = userGroup.StaffName,
                        StaffEmail = userGroup.StaffEmail,
                        Details = $"{ex.Message} \n\n Information: userEmail='{userGroup.StaffEmail}', excel rowNo={userGroup.Id} \n{ex.InnerException?.Message ?? ""}"
                    });
                    log.LogError($"FAILED: Update group for user '{userGroup.StaffEmail}'");
                }
            });
            log.LogCritical($"SUCCEEDED: Update user groups.   Ended at: {DateTime.Now}");
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

        private async Task SaveExcelToSP(ILogger log, GraphServiceClient graphClient, FunctionRunLog functionRunLog)
        {
            try
            {
                functionRunLog.LastStep = "Create excel to SP";

                string filename = $"UserGroup_{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")}.xlsx";
                using MemoryStream ms = new MemoryStream(_docContent);

                await graphClient.Drives[_documentLibrary.Id].Root.ItemWithPath(filename).Content
                    .Request()
                    .WithMaxRetry(_maxRetry)
                    .PutAsync<DriveItem>(ms);

                log.LogCritical($"SUCCEEDED: Create new excel file to SP.   Ended at: {DateTime.Now}");
            }
            catch (Exception ex)
            {
                functionRunLog.Details += $"Failed to create excel in SP. \n {ex.Message} \n";
                log.LogError($"FAILED:  create excel in SP. \n{ex.Message} \n{ex.InnerException?.Message ?? ""}");
            }
        }
    }
}
