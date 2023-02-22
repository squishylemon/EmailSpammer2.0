using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

class Program
{
    private static void Main(string[] args)
    {
        
        Console.ForegroundColor= ConsoleColor.Green;
        Console.WriteLine(@"
 /$$                                                                                            /$$$$$$       /$$$$$$ 
| $$                                                                                           /$$__  $$     /$$$_  $$
| $$        /$$$$$$$  /$$$$$$   /$$$$$$  /$$$$$$/$$$$  /$$$$$$/$$$$   /$$$$$$   /$$$$$$       |__/  \ $$    | $$$$\ $$
| $$       /$$_____/ /$$__  $$ |____  $$| $$_  $$_  $$| $$_  $$_  $$ /$$__  $$ /$$__  $$        /$$$$$$/    | $$ $$ $$
| $$      |  $$$$$$ | $$  \ $$  /$$$$$$$| $$ \ $$ \ $$| $$ \ $$ \ $$| $$$$$$$$| $$  \__/       /$$____/     | $$\ $$$$
| $$       \____  $$| $$  | $$ /$$__  $$| $$ | $$ | $$| $$ | $$ | $$| $$_____/| $$            | $$          | $$ \ $$$
| $$$$$$$$ /$$$$$$$/| $$$$$$$/|  $$$$$$$| $$ | $$ | $$| $$ | $$ | $$|  $$$$$$$| $$            | $$$$$$$$ /$$|  $$$$$$/
|________/|_______/ | $$____/  \_______/|__/ |__/ |__/|__/ |__/ |__/ \_______/|__/            |________/|__/ \______/ 
                    | $$                                                                                              
                    | $$                                                                                              
                    |__/                                                                                             
            ");
        


        Console.WriteLine("Show failed emails? (y/n)");
        string disabledFails = Console.ReadLine();
        bool showErrors = false;
        if (disabledFails.ToLower() == "y")
        {
            showErrors = true;
        }

        BaseQuestions(showErrors);
    }

    static void BaseQuestions(bool showErrors)
    {
        Console.WriteLine("Select SMTP server:");
        Console.WriteLine("1. Gmail");
        Console.WriteLine("2. Outlook");
        int serverChoice = int.Parse(Console.ReadLine());

        Console.WriteLine("Enter sender email addresses (separated by commas):");
        string fromEmails = Console.ReadLine();
        string[] fromEmailArray = fromEmails.Split(',');

        Console.WriteLine("Enter sender passwords (separated by commas):");
        string fromPasswords = Console.ReadLine();
        string[] fromPasswordArray = fromPasswords.Split(',');

        Console.WriteLine("Enter the recipient email addresses (separated by commas):");
        string toEmails = Console.ReadLine();
        string[] toEmailArray = toEmails.Split(',');

        Console.WriteLine("Enter the email subject:");
        string subject = Console.ReadLine();

        Console.WriteLine("Enter the email body:");
        string body = Console.ReadLine();

        Console.WriteLine("Do you want to attach an image file? (y/n)");
        string attachImageResponse = Console.ReadLine();

        string imagePath = null;
        if (attachImageResponse.ToLower() == "y")
        {
            Console.WriteLine("Enter the path to the image file:");
            imagePath = Console.ReadLine();
        }

        Console.WriteLine("Enter the number of emails to send:");
        int numEmails = int.Parse(Console.ReadLine());

        var tasks = new List<Task>();

        for (int i = 1; i <= numEmails; i++)
        {
            for (int j = 0; j < fromEmailArray.Length; j++)
            {
                var fromEmail = fromEmailArray[j].Trim();
                var fromPassword = fromPasswordArray[j].Trim();

                if (serverChoice == 1)
                {
                    tasks.Add(Task.Run(() => SendEmailGmail(fromEmail, fromPassword, toEmailArray, subject, body, imagePath, i, showErrors)));
                }
                else
                {
                    tasks.Add(Task.Run(() => SendEmailOutlook(fromEmail, fromPassword, toEmailArray, subject, body, imagePath, i, showErrors)));
                }
            }
        }

        Task.WhenAll(tasks).Wait();

        BaseQuestions(showErrors);
    }



    static void SendEmailOutlook(string fromEmail, string fromPassword, string[] toEmailArray, string subject, string body, string imagePath, int emailIndex, bool showErrors)
    {
        try
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("", fromEmail));
            foreach (string toEmail in toEmailArray)
            {
                message.To.Add(new MailboxAddress(toEmail.Trim(), toEmail.Trim()));
            }
            message.Subject = subject;
            var builder = new BodyBuilder();
            builder.HtmlBody = body;
            if (imagePath != null)
            {
                builder.Attachments.Add(imagePath);
            }
            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
                client.Authenticate(fromEmail, fromPassword);
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine($"Email successfully sent ({emailIndex})");
        }
        catch (Exception ex)
        {
            if (showErrors == true)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error sending email retrying...");
                Console.ForegroundColor = ConsoleColor.Green;
                SendEmailOutlook(fromEmail, fromPassword, toEmailArray, subject, body, imagePath, 0, showErrors);
            }
            SendEmailOutlook(fromEmail, fromPassword, toEmailArray, subject, body, imagePath, 0, showErrors);
        }
    }

    static void SendEmailGmail(string fromEmail, string fromPassword, string[] toEmailArray, string subject, string body, string imagePath, int emailIndex, bool showErrors)
    {
        try
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("", fromEmail));
            foreach (string toEmail in toEmailArray)
            {
                message.To.Add(new MailboxAddress(toEmail.Trim(), toEmail.Trim()));
            }
            message.Subject = subject;
            var builder = new BodyBuilder();
            builder.HtmlBody = body;
            if (imagePath != null)
            {
                builder.Attachments.Add(imagePath);
            }
            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                client.Authenticate(fromEmail, fromPassword);
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine($"Email successfully sent ({emailIndex})");
        }
        catch (Exception ex)
        {
            if (showErrors == true)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error sending email retrying...");
                Console.ForegroundColor = ConsoleColor.Green;
                SendEmailGmail(fromEmail, fromPassword, toEmailArray, subject, body, imagePath, 0, showErrors);
            }
            SendEmailGmail(fromEmail, fromPassword, toEmailArray, subject, body, imagePath, 0, showErrors);
        }
    }


}
