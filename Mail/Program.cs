using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using GemBox.Spreadsheet;
using System.Web;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using S22.Imap;
using System.Net.Mail;
using Microsoft.Vbe.Interop;
using System.Diagnostics;


namespace Mail
{
    class SystemAudit
    {
        private static Microsoft.Office.Interop.Excel.Application SysAuditXLWApp;
        private static Microsoft.Office.Interop.Excel.Application SysAuditXLWApp2;
        private static Microsoft.Office.Interop.Excel.Workbook SysAuditWWorkBook;
        private static Microsoft.Office.Interop.Excel.Workbook SysAuditWWorkBook2;
        private static Microsoft.Office.Interop.Excel.Worksheet SysAuditWWorkSheet;
        private static Microsoft.Office.Interop.Excel.Worksheet SysAuditWWorkSheet2;

        static void Main(string[] args)
        {

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("################################################################################");
            Console.WriteLine("#####################            WinAuditPro          ##########################");
            Console.WriteLine("################################################################################");
            Console.WriteLine();


            sendCompletionNotification("brodriguez@statesidebpo.com");




        }

        private static void sendCompletionNotification(string email)
        {
            try
            {


                SmtpClient mySmtpClient = new SmtpClient("secure.emailsrvr.com", 25);
                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new
                System.Net.NetworkCredential("notify@statesidebpo.com", "W31is+en2016");
                mySmtpClient.Credentials = basicAuthenticationInfo;

                // add from,to mailaddresses
                MailAddress from = new MailAddress("notify@statesidebpo.com");
                MailAddress to = new MailAddress("helpdesk@statesidebpo.com"); //to = new MailAddress("brodriguez@statesidebpo.com"); //
                MailMessage myMail = new MailMessage(from, to);
                myMail.Subject = DateTime.Now + " SystemAudit Processing run completed successfully";

            
                myMail.IsBodyHtml = true;

                string bd = "";
                if (email.Count() == 1)
                {
                    bd = "<h2> " + email.Count() + " Audit was processed</h2>";
                }
                else
                {
                    bd = "<h2> " + email.Count() + " Audits were processed</h2>";
                }

                if (email.Count(!= 0))
                {

                    bd = bd + "<table style = " + "' tr:nth-child(even) {background-color: #f2f2f2} '" + " border = " + "1" + " border-radius= " + "10px" + " cellpadding = " + "6" + " cellspacing = " + "5" + "><tbody>";


                    foreach (SysAuditResults c in CandidatesList)
                    {
                        string reason = c.aFailedReason;

                        if (c.aFailedReason != null)
                        {
                            reason = " - " + c.aFailedReason;
                        }

                        if (c.needsManualProcessing)
                        {
                            bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'33%'" + ">" + c.cEmail + "</td><td width=" + "'43%'" + ">" + c.aResult + reason + "</td></tr>";
                            sendMail(c.cEmail, c.attachmentFilename, c.cName);
                        }
                        else
                        {
                            if (c.aResult == "Fail")
                            {
                                bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'43%'" + ">" + c.cEmail + "</td><td  width=" + "'43%'" + "><font color='red'>" + c.aResult + reason + "</font></td></tr>";
                            }
                            if (c.aResult == "Pass")
                            {
                                bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'28%'" + ">" + c.cEmail + "</td><td ' width=" + "'43%'" + "><font color='green'>" + c.aResult + reason + "</font></td></tr>";
                            }

                            if (c.aResult == "Pending")
                            {
                                bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'28%'" + ">" + c.cEmail + "</td><td  width=" + "'43%'" + "><font color='blue'>" + c.aResult + reason + "</font></td></tr>";
                            }

                            sendMail(c.cEmail, c.attachmentFilename, c.cName);
                        }

                    }

                    bd = bd + "</tbody></table>";
                }
                myMail.Body = Regex.Replace(bd, @"[^\u0000-\u007F]", " ");
                mySmtpClient.Send(myMail);

            }
            catch (System.Exception ex)
            {
                throw new ApplicationException
                  ("Outlook exception has occured: " + ex.Message);
            }
        }

        public static void sendMail(string recipient, string attachmentFilename, string cadidateName)
        {
            try
            {
                bool ex = File.Exists(attachmentFilename);

                if (ex)
                {
                    SmtpClient mySmtpClient = new SmtpClient("secure.emailsrvr.com", 25);
                    mySmtpClient.UseDefaultCredentials = false;
                    System.Net.NetworkCredential basicAuthenticationInfo = new
                    System.Net.NetworkCredential("notify@statesidebpo.com", "W31is+en2016");
                    mySmtpClient.Credentials = basicAuthenticationInfo;

                    MailAddress from = new MailAddress("notify@statesidebpo.com");
                    MailAddress to = new MailAddress(recipient);
                    MailAddress cc = new MailAddress("recruiters@statesidebpo.com");

                  

                    MailMessage myMail = new MailMessage(from, to);
                    myMail.IsBodyHtml = true;
                    myMail.Subject = "System audit results for " + cadidateName;
                    myMail.CC.Add(cc);

                    if (isTESTING == true)
                    {
                        myMail.Subject = "TESTING - System audit results for " + cadidateName;
                    }

                    string body = @"<p style =""font-size:21px"">Dear " + cadidateName + ",<br><br>" + "This email is to inform you of your system audit results. Please see the attachment. If you have any technical questions regarding your results, please reach out to us via email at <a mailto:winaudit@statesidebpo.com>winaudit@statesidebpo.com</a>.</p>";
                    body = body + @"<p style =""font-size:18px""><i>(Please note: If you are unable to view the attachment, you may need to download and install Adobe Acrobat Reader DC or a similar program that allows the viewing of PDF documents)</i>";

                    Attachment inlineStatetSideLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\StatesideLogo.png");
                    Attachment inlineBitLeverLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\Bit-LeverLogo.png");
                    Attachment SysAuditResults = new Attachment(@"\\Filesvr4\IT\WinAudit\Results_Archive\" + cadidateName + " SystemAudit Results.pdf");


                    if (isTESTING == true)
                    {
                        SysAuditResults = new Attachment(@"\\Filesvr4\it\WinAudit\Results_Archive\" + "TESTING - " + cadidateName + " SystemAudit Results.pdf");
                    }

                    myMail.Attachments.Add(SysAuditResults);
                    myMail.Attachments.Add(inlineStatetSideLogo);
                    myMail.Attachments.Add(inlineBitLeverLogo);

                    inlineBitLeverLogo.ContentDisposition.Inline = true;
                    inlineStatetSideLogo.ContentDisposition.Inline = true;

                    inlineBitLeverLogo.ContentId = "BitLeverLogo";
                    inlineStatetSideLogo.ContentId = " StatetSideLogo";
                    SysAuditResults.ContentId = "Pdf";


                    inlineBitLeverLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline;
                    inlineStatetSideLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline;

                    myMail.Body = @"<htm><body>" + body + "<br>  <table><tr><td><img src=\"cid:StatetSideLogo\"></td><td><img src=\"cid:BitLeverLogo\"></td></tr></table></body></html>";
                    myMail.BodyEncoding = System.Text.Encoding.UTF8;

                    mySmtpClient.Send(myMail);

                }

            }
            catch (SmtpException ex)
            {
                throw new ApplicationException
                  ("SmtpException has occured: " + ex.Message);
            }


        }
    }
}
