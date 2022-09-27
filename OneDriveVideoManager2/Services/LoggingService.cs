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

        public static async Task CreateFunctionRunLogToSP(
            ILogger log,
            GraphServiceClient graphClient,
            FunctionRunLog functionRunLog
        )
        {
            try
            {
                if (functionRunLog.Status == "Running")
                {
                    functionRunLog.Status = "Succeeded";
                }
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");
                string targetListName = Environment.GetEnvironmentVariable("APPSETTING_FunctionRunLogListName");

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

                log.LogCritical($"SUCCEEDED: Create SP functionRunLog. Time: {DateTime.Now}");
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
                string hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
                string spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");
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
                    await graphClient.Sites.GetByPath(spSiteRelativePath, hostName).Lists[targetListName].Items
                        .Request()
                        .WithMaxRetry(_maxRetry)
                        .AddAsync(newItem);
                });
                
                log.LogCritical($"SUCCEEDED: Create SP errorLog(s). Time: {DateTime.Now}");
            }
            catch (Exception ex)
            {
                functionRunLog.Details += $"{ex.Message} \n{ex.InnerException?.Message ?? ""}";
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
            }
        }
    }
}
