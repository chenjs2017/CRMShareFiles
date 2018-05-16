using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Utility
{
    public class EmailTool
    {
         // should turn off secure https://support.google.com/accounts/answer/6010255?hl=en-GB
        public static void SendEmailFromGmail(string toEmail, string toName ,string ccEmail, string ccName, string subject, string body,string fromEmail, string username, string password)
        {
            //            var fromAddress = new MailAddress("zimmetsummit@gmail.com", "Zimmet");
            var fromAddress = new MailAddress(fromEmail);
            var toAddress = new MailAddress(toEmail,toName);


 //           const string userName = "zimmetsummit@gmail.com";
  //          const string fromPassword = "Dynamic2018";

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(username , password)
            };
            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                if (ccEmail != null && ccName != null)
                {
                    var ccAddress = new MailAddress(ccEmail, ccName);
                    message.CC.Add(ccAddress);
                }
                smtp.Send(message);
            }
        }
    }
}
