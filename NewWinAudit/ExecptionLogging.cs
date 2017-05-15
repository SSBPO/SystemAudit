using System;
using System.Net.Mail;
using System.Configuration;
using System.Net;
using context = System.Web.HttpContext;

namespace NewWinAudit
{

public static class ExceptionLoggingClass
    {

        private static String ErrorlineNo, Errormsg, ErrorLocation, extype, exurl, Frommail, ToMail, Sub, HostAdd, EmailHead, EmailSing;


        public static void SendErrorTomail(Exception exmail)
        {
            string ErrorlineNo, Errormsg, ErrorLocation, extype, exurl, Frommail, ToMail, Sub, HostAdd, EmailHead, EmailSing;

            try
            {
                var newline = "<br/>";
                ErrorlineNo = exmail.StackTrace.Substring(exmail.StackTrace.Length - 7, 7);
                Errormsg = exmail.GetType().Name.ToString();
                extype = exmail.GetType().ToString();
                exurl = context.Current.Request.Url.ToString();
                ErrorLocation = exmail.Message.ToString();

                EmailHead = "<b>Dear Team,</b>" + "<br/>" + "An exception occurred in a Application Url" + " " + exurl + " " + "With following Details" + "<br/>" + "<br/>";
                EmailSing = newline + "Thanks and Regards" + newline + "    " + "     " + "<b>Application Admin </b>" + "</br>";
                Sub = "Exception occurred" + " " + "in Application" + " " + exurl;
               // HostAdd = ConfigurationManager.AppSettings["Host"].ToString();
                string errortomail = EmailHead + "<b>Log Written Date: </b>" + " " + DateTime.Now.ToString() + newline + "<b>Error Line No :</b>" + " " + ErrorlineNo + "\t\n" + " " + newline + "<b>Error Message:</b>" + " " + Errormsg + newline + "<b>Exception Type:</b>" + " " + extype + newline + "<b> Error Details :</b>" + " " + ErrorLocation + newline + "<b>Error Page Url:</b>" + " " + exurl + newline + newline + newline + newline + EmailSing;

                using (MailMessage mailMessage = new MailMessage())
                {
                    Frommail = "helpdesk@statesidebpo.com";
                    ToMail = "brodriguez@statesidebpo.com";
      

                    mailMessage.From = new MailAddress(Frommail);
                    mailMessage.Subject = Sub;
                    mailMessage.Body = errortomail;
                    mailMessage.IsBodyHtml = true;

                    string[] MultiEmailId = ToMail.Split(',');
                    foreach (string userEmails in MultiEmailId)
                    {
                        mailMessage.To.Add(new MailAddress(userEmails));
                    }


                    SmtpClient mySmtpClient = new SmtpClient("secure.emailsrvr.com", 25);
                    mySmtpClient.UseDefaultCredentials = false;
                    System.Net.NetworkCredential basicAuthenticationInfo = new
                    System.Net.NetworkCredential("systemaudit@statesidebpo.com", "Uhtd5$#8s776fsdfa4!df!!2eX");
                    mySmtpClient.Credentials = basicAuthenticationInfo;


                    mySmtpClient.Send(mailMessage); //sending Email  

                }
            }
            catch (Exception em)
            {
                em.ToString();

            }
        }

    }
}
