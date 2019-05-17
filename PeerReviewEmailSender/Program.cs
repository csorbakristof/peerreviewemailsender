using System;

namespace PeerReviewEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            var p = new Program();
            p.Doit();
        }

        private async void Doit()
        {
            Console.WriteLine("Password for email account: ");

            ConsoleColor origBG = Console.BackgroundColor;
            ConsoleColor origFG = Console.ForegroundColor;
            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.Red;
            string password = Console.ReadLine();
            Console.BackgroundColor = origBG;
            Console.ForegroundColor = origFG;

            const bool reallySendEmails = true;
            var mailer = new EmailSender("kristof@aut.bme.hu", password, reallySendEmails);
            var container = new EmailContainer();
            var reader = new FeedbackReader(container);

            Console.WriteLine("Initializing Exchange connection...");

            await mailer.SendInitialEmail();

            // Meanwhile...
            const string xlsFilename = @"e:\temp\Beszámoló értékelés (Responses).xlsx";
            reader.ReadFeedbacks(xlsFilename);

            // After initial email was sent...
            foreach (var presenter in container.GetPresenterEmails())
            {
                var body = container.GetMessageBody(presenter);
                var subject = "Beszámoló visszajelzések";
                await mailer.SendEmailAsync(presenter, subject, body);
            }

            Console.WriteLine("All feedback emails have been sent. Enter something and press Enter...");
            Console.ReadLine();
        }
    }
}
