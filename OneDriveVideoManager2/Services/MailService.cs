using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Extensions.Logging;
using OneDriveVideoManager.Models;
using Azure.Identity;
using static System.Formats.Asn1.AsnWriter;

namespace OneDriveVideoManager.Services
{
    public static class MailService
    {
        private static readonly bool _useTestTenant = bool.Parse(Environment.GetEnvironmentVariable("APPSETTING_SendEmailWithTestTenant"));
        private static readonly string _environment = (_useTestTenant) ? "-Test": "";
        private static readonly string _emailSender = Environment.GetEnvironmentVariable($"APPSETTING_EmailSender{_environment}");
        private static readonly string _emailRecipient = Environment.GetEnvironmentVariable($"APPSETTING_EmailRecipient{_environment}");
        private static readonly string _emailCCRecipient = Environment.GetEnvironmentVariable($"APPSETTING_EmailCCRecipient{_environment}");
        private static readonly string _functionAppName = Environment.GetEnvironmentVariable("APPSETTING_FunctionAppName");
        private static readonly string _spLogListURL = Environment.GetEnvironmentVariable("APPSETTING_FunctionRunLogListURL");
        private static readonly string _githubURL = Environment.GetEnvironmentVariable("APPSETTING_GithubURL");
        private static readonly string _hostName = Environment.GetEnvironmentVariable("APPSETTING_HostName");
        private static readonly string _spSiteRelativePath = Environment.GetEnvironmentVariable("APPSETTING_SpSiteRelativePath");
        private static readonly int _maxRetries = 3;

        
        public static async Task SendNotificationEmail(ILogger log, GraphServiceClient graphClient, Checker checker)
        {
            try
            {
                string emailSubject = "New batch of videos shared for checking";
                string emailBodyContent = $"Hi Checkers, <br /><br /> <b>{checker.Videos.Count} new videos</b> are shared to you, please check using the links below and update checking status on Sharepoint. <br /><br />";
                emailBodyContent += $"SharePoint list: https://{_hostName}/{_spSiteRelativePath}/Lists/{checker.ListName}  <br />";
                emailBodyContent += $"Video(s): <br /><ol>";
                foreach (Recording recording in checker.Videos)
                {
                    emailBodyContent += $"<li> <a href='{recording.Link}'>{recording.Name}</a>, duration: {recording.Duration} </li>";
                }
                emailBodyContent += "</ol><br /> Thanks, <br /><br />";
                emailBodyContent += "Automated function";

                ItemBody emailBody = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = emailBodyContent,
                };
                List<Recipient> emailRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = (_useTestTenant) ? _emailRecipient : checker.Email
                        }
                    }
                };
                Message message = new Message
                {
                    Subject = emailSubject,
                    Body = emailBody,
                    ToRecipients = emailRecipients
                };

                await graphClient.Users[_emailSender]
                    .SendMail(message, null)
                    .Request()
                    .WithMaxRetry(_maxRetries)
                    .PostAsync();

                log.LogInformation($"> Sent notification email to checker: {checker.Email}.   Time: {DateTime.Now}");
            }
            catch (Exception ex)
            {
                log.LogError($"Unable to email checker: {checker.Email} \n{ex.Message} \n{ex.InnerException?.Message ?? ""}");
                throw;
            }
        }

        public static async Task SendReportErrorEmail(ILogger log, GraphServiceClient graphClient, string functionName, string errorMsg)
        {
            try
            {
                string datetime = DateTime.Now.ToString("yyyy-MM-dd");
                string emailSubject = $"Failed case(s) detected when running Azure-function: {functionName}";
                string emailBodyContent = $"Dear Developers,  <br /><br /> <b>Failed case(s)</b> of {functionName} function are detected on {datetime}. Details as below: <br /><br />";
                emailBodyContent += $"FunctionApp: {_functionAppName} <br />";
                emailBodyContent += $"SPLogList: {_spLogListURL} <br />";
                emailBodyContent += $"Github: {_githubURL} <br /><br />";
                emailBodyContent += $"<p> Status: Failed <br /> Details: {errorMsg} </p> <br />";
                emailBodyContent += "<p> Please check. </p>";

                ItemBody emailBody = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = emailBodyContent,
                };
                List<Recipient> emailRecipients = new List<Recipient>();
                List<Recipient> emailCCRecipients = new List<Recipient>();
                string[] allRecipients = _emailRecipient.Split(';');
                string[] allCCRecipients = _emailCCRecipient.Split(';');

                foreach (string recipient in allRecipients)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        emailRecipients.Add(new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = recipient
                            }
                        });
                    }
                }
                foreach (string recipient in allCCRecipients)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        emailCCRecipients.Add(new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = recipient
                            }
                        });
                    }
                }
                Message message = new Message
                {
                    Subject = emailSubject,
                    Body = emailBody,
                    ToRecipients = emailRecipients,
                    CcRecipients = emailCCRecipients
                };

                await graphClient.Users[_emailSender]
                    .SendMail(message, null)
                    .Request()
                    .WithMaxRetry(_maxRetries)
                    .PostAsync();

                log.LogCritical($"SUCCEEDED: Send error reporting email to {_emailRecipient}.   Time: {DateTime.Now}");

            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
            }
        }
    }
}
