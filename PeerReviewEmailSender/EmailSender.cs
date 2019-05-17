using Microsoft.Exchange.WebServices.Data;
using System;
using System.Threading.Tasks;

namespace PeerReviewEmailSender
{
    public class EmailSender
    {
        private readonly string user;
        private readonly string password;
        private readonly bool reallySendEmails;
        public EmailSender(string user, string password, bool reallySendEmails=true)
        {
            this.user = user;
            this.password = password;
            this.reallySendEmails = reallySendEmails;
        }

        public async System.Threading.Tasks.Task SendInitialEmail()
        {
            if (reallySendEmails)
                await SendEmailAsync(user, "Initial email", "This e-mail is sent to initialize Exchange Web Services...");
        }

        public async System.Threading.Tasks.Task SendEmailAsync(string toEmail, string subject, string body)
        {
            Console.WriteLine($"Sending email to {toEmail} (subject: {subject}):\n{body}\n\n");
            if (!reallySendEmails)
                return;

            var exchange = await GetExchangeServiceAsync();
            var msg = new EmailMessage(exchange)
            {
                From = new EmailAddress(this.user),
                Subject = subject,
                Body = body
            };

            msg.ToRecipients.Add(new EmailAddress(toEmail));
            msg.SendAndSaveCopy(new FolderId(WellKnownFolderName.SentItems));
        }

        private ExchangeService exchangeService;
        private async Task<ExchangeService> GetExchangeServiceAsync()
        {
            if (exchangeService == null)
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
                {
                    Credentials = new WebCredentials(user, password, "AUT"),
                    TraceEnabled = true,
                    TraceFlags = TraceFlags.All
                };
                service.AutodiscoverUrl(user, redirectionUri => new Uri(redirectionUri).Scheme == "https");
                exchangeService = service;
            }
            return exchangeService;
        }


    }
}
