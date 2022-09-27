using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Extensions.Logging;

namespace OneDriveVideoManager.Services
{
    public static class MailService
    {
        private static readonly string _emailSender = Environment.GetEnvironmentVariable("APPSETTING_EmailSender");
        private static readonly string _emailRecipient = Environment.GetEnvironmentVariable("APPSETTING_EmailRecipient");
        private static readonly string _emailCCRecipient = Environment.GetEnvironmentVariable("APPSETTING_EmailCCRecipient");
        private static readonly string _functionAppName = Environment.GetEnvironmentVariable("APPSETTING_FunctionAppName");
        private static readonly string _spLogListURL = Environment.GetEnvironmentVariable("APPSETTING_FunctionRunLogListURL");
        private static readonly string _githubURL = Environment.GetEnvironmentVariable("APPSETTING_GithubURL");
        private static readonly int _maxRetries = 3;

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
                string[] testRecipients = _emailRecipient.Split(';');
                string[] testCCRecipients = _emailCCRecipient.Split(';');

                foreach (string recipient in testRecipients)
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
                foreach (string recipient in testCCRecipients)
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

                log.LogCritical($"SUCCEEDED: Send error reporting email to {_emailRecipient}. Time: {DateTime.Now}");

            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message} \n{ex.InnerException?.Message ?? ""}");
            }
        }
    }
}
