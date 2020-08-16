using Giveaway.Helper;
using Giveaway.Infra;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using SendGrid;
using SendGrid.Helpers.Mail;
using System.Threading.Tasks;
using System.Net.Mail;

namespace Giveaway.Services
{
    public class EmailSendGrid: IEmailSendGrid
    {
        private  IConfiguration configuration { get; set; }
        private readonly AppSettings settings;
        public EmailSendGrid(IConfiguration _configuration,IOptions<AppSettings> _options)
        {
            configuration = _configuration;
            settings = _options.Value;
        }


        public Task SendEmailAsync(string email, string subject, string message)
        {
            return Execute(email,subject, message);
        }

        public Task Execute(string email, string subject, string message)
        {

            //Tested sending email through gmail SMTP server and it is working for my email id and password. 
            //Before that we have to go to "https://myaccount.google.com/intro/security -> Less secure app access -> Turn on Access" 

            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress("NO_REPLY@noreply.com.au");
            mailMessage.To.Add(new MailAddress(email));
            mailMessage.Subject = subject;
            mailMessage.IsBodyHtml = true;
            mailMessage.Body = message;
            SmtpClient client = new SmtpClient();
            client.Port = 587;
            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(settings.EmailUsername, settings.EmailPassword);
            client.Host = "smtp.gmail.com";
            return client.SendMailAsync(mailMessage);

            //Tried Send grid with api key, it worked initially then my email is diagnosed as spammer and blocked so coded back for SMTP.
            //var client = new SendGridClient(apiKey);
            //var msg1 = MailHelper.CreateSingleEmail(new EmailAddress("manu_net@outlook.com.au", "manu"),new EmailAddress("manu_net@outlook.com.au", "manu") , "Test subject", "", "message");
            //return client.SendEmailAsync(msg1);
        }
    }
}
