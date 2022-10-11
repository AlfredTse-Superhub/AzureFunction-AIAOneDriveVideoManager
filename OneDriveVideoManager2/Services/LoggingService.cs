using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using OneDriveVideoManager.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDriveVideoManager.Services
{
    public static class LoggingService
    {
        private static readonly int _maxRetry = 2;
        private static readonly string _hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
        private static readonly string _spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");
        private static readonly string _targetListName = Environment.GetEnvironmentVariable("APPSETTING_FunctionRunLogListName");

        public static async Task PostFunctionRunLogToSP(
            ILogger log,
            GraphServiceClient graphClient,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                if (functionRunLog.Status != "Failed")
                {
                    functionRunLog.Status = "Succeeded";
                }

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
                if (functionRunLog.FunctionName.ToLower() == "updateusergroup")
                {
                    newItem.Fields.AdditionalData.Add("TotalRecords", functionRunLog.TotalRecords);
                    newItem.Fields.AdditionalData.Add("UpdatedRecords", functionRunLog.UpdatedRecords);
                }

                if (string.IsNullOrWhiteSpace(functionRunLog.Id))
                {
                    ListItem newLog = await graphClient.Sites.GetByPath(_spSiteRelativePath, _hostName).Lists[_targetListName].Items
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .AddAsync(newItem); // create new

                    functionRunLog.Id = newLog.Id;
                }
                else
                {
                    await graphClient.Sites.GetByPath(_spSiteRelativePath, _hostName).Lists[_targetListName].Items[functionRunLog.Id]
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .UpdateAsync(newItem); // update
                }

                log.LogCritical($"SUCCEEDED: Post SP functionRunLog.   Time: {DateTime.Now}");

            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
            }
        }

        public static async Task CreateErrorLogToSP(
            ILogger log,
            GraphServiceClient graphClient,
            string functionName,
            string targetListName,
            List<ErrorLog> errorLogs,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                functionRunLog.LastStep = "Create ErrorLog(s) to SP";

                await Parallel.ForEachAsync(errorLogs, async (errorLog, cancellationToken) => 
                {
                    ListItem newItem = new ListItem
                    {
                        Fields = new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object>()
                            {
                                {"FunctionName", functionName},
                                {"StaffName", errorLog.StaffName},
                                {"StaffEmail", errorLog.StaffEmail},
                                {"Details", errorLog.Details}
                            }
                        }
                    };
                    await graphClient.Sites.GetByPath(_spSiteRelativePath, _hostName).Lists[targetListName].Items
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .AddAsync(newItem);
                });
                
                log.LogCritical($"SUCCEEDED: Create SP errorLog(s).   Time: {DateTime.Now}");

            }
            catch (Exception ex)
            {
                functionRunLog.Details += $"{ex.Message} \n{ex.InnerException?.Message ?? ""}";
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
            }
        }
    }
}
