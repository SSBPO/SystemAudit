using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Web;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using S22.Imap;
using GemBox.Spreadsheet;
using System.Reflection;
using System.Net.Mail;
using System.IO;
using System.Net.Mime;
using System.Threading;

namespace NewWinAudit
{
    class SystemAudit
    {       
        static void Main(string[] args)
        {
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("################################################################################");
            Console.WriteLine("#####################            WinAuditPro          ##########################");
            Console.WriteLine("################################################################################");
            Console.WriteLine();

            // Declare app instance / list to hold all sysaudits 
            SystemAudit xxx_SysAudit = new SystemAudit();
            List<SysAuditResults> CandidatesList = new List<SysAuditResults>();

            string excelLicense = "EQU2-1K6F-UZPP-4MOX";
            SpreadsheetInfo.SetLicense(excelLicense);

            using (S22.Imap.ImapClient imapClient = new S22.Imap.ImapClient("secure.emailsrvr.com", 993, "systemaudit@statesidebpo.com", "Stateside@2017", AuthMethod.Login, true))
            {
                // Strip HTML from email class
                killExcel();
                HtmlToText stripHtml = new HtmlToText();

                // Check mailbox and get any messages not seen and sent by systemaudit@bit-lever.com
                IEnumerable<uint> uids = imapClient.Search(S22.Imap.SearchCondition.Unseen().And(S22.Imap.SearchCondition.From("systemaudit@bit-lever.com")));
                IEnumerable<System.Net.Mail.MailMessage> messages = imapClient.GetMessages(uids);

                // Start program
                int ProcessedEmails = 0;
                Console.WriteLine("Checking SystemAudits mailbox.");

                // If there are new unseen messages 
                if (messages.Count() >= 1)
                {
                    Console.WriteLine("There is " + messages.Count() + " SystemAudit to process.");
                    // Send them to process 
                    ProcessedEmails = sysAuditProcess(CandidatesList, stripHtml, messages, ProcessedEmails);
                }
                else
                {   // show there are 0 messages 
                    Console.WriteLine("There are " + messages.Count() + " SystemAudits to process.");
                }

                // End processing and say bye 
                Console.WriteLine(ProcessedEmails + " email(s) processed. Good bye!");
                Thread.Sleep(5000);
            }
        }
        
        // Function to process system audit messages
        private static int sysAuditProcess(List<SysAuditResults> CandidatesList, HtmlToText stripHtml, IEnumerable<MailMessage> messages, int ProcessedEmails)
        {
            if (messages.Count() != 0)
            {
                Console.WriteLine("");
                Console.WriteLine("Starting...");
                Console.WriteLine("");

                // For each message
                foreach (System.Net.Mail.MailMessage msg in messages)
                {
                    //Load the WinAudit workbook
                    string fileName = @"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xlsm";
                    var ef = ExcelFile.Load(fileName);

                    //Instanciate SystemAudit Results object
                    SysAuditResults SSBPOsysAuditResults = new SysAuditResults();

                    // If message body is not empty 
                    if (msg.Body != "")
                    {
                        //Strip Html from emal and split into 14 lines  
                        string[] line = stripHtml.Convert(msg.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);                            

                        Console.WriteLine("Processing " + ProcessedEmails + " of " + messages.Count() + ".");

                        //Foreach of the 14 lines in the email
                        foreach (string l in line.ToList()) 
                        {
                            if (l.ToString() != "" & !l.ToString().Contains("Simplified Audit Results") & !l.ToString().Contains("www")) //If the line is not empty or has unwanted text
                            {
                                //Add the values to the the SystemAudit results object
                                SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, l);
                            }
                        }
                    }

                    //Fill out the candidates results worksheet and save as pdf to send out

                    ExcelWorksheet worksheet = ef.Worksheets.ActiveWorksheet;
                    worksheet.Cells["C5"].Value = SSBPOsysAuditResults.cName;
                    worksheet.Cells["C6"].Value = SSBPOsysAuditResults.auditDate;
                    worksheet.Cells["D6"].Value = SSBPOsysAuditResults.aResult;
                    if (SSBPOsysAuditResults.aResult == "Passed")
                    {
                        worksheet.Cells["D6"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D6"].Style.Font.Color = System.Drawing.Color.Red;
                    }
                    worksheet.Cells["C13"].Value = SSBPOsysAuditResults.cOS.Substring(0, 10).Replace(".", "").Trim(); 
                    worksheet.Cells["C14"].Value = SSBPOsysAuditResults.cHDD.Substring(64, 6).Replace("]", "").Trim(); 
                    worksheet.Cells["C15"].Value = SSBPOsysAuditResults.cRAM.Substring(33, 4).Replace("]", " ") + " RAM".Trim(); 
                    worksheet.Cells["C16"].Value = SSBPOsysAuditResults.cCPU.Substring(11, 4).Trim();
                    worksheet.Cells["C18"].Value = SSBPOsysAuditResults.cInternetUp.Trim();
                    worksheet.Cells["C19"].Value = SSBPOsysAuditResults.cInternetDown.Trim();
                    worksheet.Cells["D13"].Value = SSBPOsysAuditResults.OSaResult;
                    worksheet.Cells["D14"].Value = SSBPOsysAuditResults.HDDaResult;
                    worksheet.Cells["D15"].Value = SSBPOsysAuditResults.RAMaResult;
                    worksheet.Cells["D16"].Value = SSBPOsysAuditResults.CPUaResult;
                    worksheet.Cells["D18"].Value = SSBPOsysAuditResults.InternetUpResult;
                    worksheet.Cells["D19"].Value = SSBPOsysAuditResults.InternetDownResult;
                    if (SSBPOsysAuditResults.OSaResult == "Y")
                    {
                        worksheet.Cells["D13"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D13"].Style.Font.Color = System.Drawing.Color.Red;
                    }
                    if (SSBPOsysAuditResults.HDDaResult == "Y")
                    {
                        worksheet.Cells["D14"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D14"].Style.Font.Color = System.Drawing.Color.Red;
                    }
                    if (SSBPOsysAuditResults.RAMaResult == "Y")
                    {
                        worksheet.Cells["D15"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D15"].Style.Font.Color = System.Drawing.Color.Red;
                    }
                    if (SSBPOsysAuditResults.CPUaResult == "Y")
                    {
                        worksheet.Cells["D16"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D16"].Style.Font.Color = System.Drawing.Color.Red;
                    }
                    if (SSBPOsysAuditResults.InternetUpResult == "Y")
                    {
                        worksheet.Cells["D18"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D18"].Style.Font.Color = System.Drawing.Color.Red;
                    }
                    if (SSBPOsysAuditResults.InternetDownResult == "Y")
                    {
                        worksheet.Cells["D19"].Style.Font.Color = System.Drawing.Color.Green;
                    }
                    else
                    {
                        worksheet.Cells["D19"].Style.Font.Color = System.Drawing.Color.Red;
                    }

                    //Save results as pdf
                    ef.Save(@"\\Filesvr4\it\WinAudit\Results_Archive\" + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf");
                    ef = null;
                    // Add the completed SysAudit results object to the list of results
                    CandidatesList.Add(SSBPOsysAuditResults); 
                    ProcessedEmails = ProcessedEmails + 1;
                }

                createBitLeverImport(CandidatesList);
                sendResults(CandidatesList);
                sendCompletionNotification(CandidatesList);



            }
            else
            {
                Console.WriteLine("There is nothing to process. Good bye!");
            }

            return ProcessedEmails;
        }
        // Function create workbook to import results into Bit-Lever
        private static void createBitLeverImport(List<SysAuditResults> CandidatesList)
        {
                //Generate a unique name for the import spreadsheet
                string file2Import = string.Format(@"\\filesvr4\IT\WinAudit\4BitLeverImport\BitLeverImport{0:yyyy-MM-dd_hh-mm-ss-tt}" + " Results.xls", DateTime.Now);
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Bit-Lever Import");
                int index = 1;

                //Write the headers
                worksheet.Cells["A1"].Value = "Date Processed";
                worksheet.Cells["B1"].Value = "Candidate Name";
                worksheet.Cells["C1"].Value = "Candidate Email";
                worksheet.Cells["D1"].Value = "Notes";
                worksheet.Cells["E1"].Value = "SysAudit Status";
                worksheet.Cells["F1"].Value = "Processed By";
                worksheet.Cells["G1"].Value = "Results Sent?";
                worksheet.Cells["H1"].Value = "Fail Reason";

                //write a row for each systemaudit result in the list
                foreach (SysAuditResults r in CandidatesList)
                {
                    worksheet.Cells[index, 0].Value = r.auditDate;
                    worksheet.Cells[index, 1].Value = r.cName;
                    worksheet.Cells[index, 2].Value = r.cEmail;
                    worksheet.Cells[index, 3].Value = r.aResultSummary;
                    worksheet.Cells[index, 4].Value = r.aResult;
                    worksheet.Cells[index, 5].Value = Environment.UserName;
                    worksheet.Cells[index, 6].Value = "Yes";
                    worksheet.Cells[index, 7].Value = r.aFailedReason;
                    index = index + 1;
                }

                workbook.Save(file2Import);
           

        }

        // Function to send results to candidates or helpdesk if the process fails
        private static void sendResults(List<SysAuditResults> CandidatesList)
        {
            // For each candidate results
            foreach (SysAuditResults r in CandidatesList)
            {
                //Check if there is a pdf file with the candidate's name
                string attachmentFilename = @"\\filesvr4\IT\WinAudit\Results_Archive\" + r.cName.Trim() + " SystemAudit Results.pdf";
                if (File.Exists(@"\\filesvr4\IT\WinAudit\Results_Archive\" + r.cName.Trim() + " SystemAudit Results.pdf"))
                {
                    //If there is, then write results, send email to candidate 
                    Console.WriteLine(r.cName + " - " + r.aResult + ". Sending results...");
                    sendMail("brodriguez@statesidebpo.com", attachmentFilename, r.cName, "brodriguez@statesidebpo.com");
                    Console.WriteLine(r.cName + " results were sent.");
                    Console.WriteLine("");

                }
                else
                {
                    //If not then send email to helpdesk
                    sendErrorMail("brodriguez@statesidebpo.com", r.cName);
                }
            }

        }
        // Function to get the values from the email lines and populate the SyAudit Result object
        private static SysAuditResults getResultsObject(SysAuditResults sysAuditResults, string l)
        {
            if (l.Contains("Time"))
            {
                sysAuditResults.auditDate = l.ToString().Substring(17);
            }

            if (l.Contains("Full"))
            {
                sysAuditResults.cName = l.ToString().Substring(11);
            }

            if (l.Contains("Email"))
            {
                sysAuditResults.cEmail = l.ToString().Substring(15);
            }

            if (l.Contains("Hard Drive"))
            {
                sysAuditResults.cHDD = l.ToString().Substring(27);

                string[] cHDDs = sysAuditResults.cHDD.Split('=');
                string cHDD1 = cHDDs[3].Replace("GB]", "");

                if (Convert.ToInt32(cHDD1) < 25)
                {
                    sysAuditResults.HDDaResult = "F";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", " + "Hard drive insufficient - Hardware / System issues box";
                }
                else
                {
                    sysAuditResults.HDDaResult = "Y";
                }
            }

            if (l.Contains("Operating"))
            {
                sysAuditResults.cOS = l.ToString().Substring(18);

                string[] cOSs = sysAuditResults.cOS.Split(' ');
                string cOS1 = cOSs[1];
                if (cOS1.Length > 1)
                {
                    cOSs = cOS1.Split('.');
                    cOS1 = cOSs[0];
                }

                if (Convert.ToInt64(cOS1) < 7)
                {
                    sysAuditResults.OSaResult = "N";
                    sysAuditResults.aFailedReason = "OS insufficient - Hardware / System issues box";
                }
                else
                {
                    sysAuditResults.OSaResult = "Y";
                }

            }
            if (l.Contains("Computer"))
            {
                sysAuditResults.cHost = l.ToString().Substring(15);
            }

            if (l.Contains("RAM"))
            {
                sysAuditResults.cRAM = l.ToString().Substring(22);

                sysAuditResults.cRAM = l.ToString().Substring(22);
                string[] cRAMs = sysAuditResults.cRAM.Split('=');
                string cRAM1 = cRAMs[2].Replace("GB]", "");

                if (Convert.ToInt32(cRAM1) < 2)
                {
                    sysAuditResults.RAMaResult = "N";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "; " + "RAM insufficient - Hardware / System issues box";
                }
                else
                {
                    sysAuditResults.RAMaResult = "Y";
                }
            }

            if (l.Contains("Network"))
            {
                //" 905.46 Kbps]"

                sysAuditResults.cInternetUp = l.ToString().Substring(62).Replace("]", "");

                if (sysAuditResults.cInternetUp.Contains("Kbps"))
                {


                    if (sysAuditResults.cInternetUp[0] < 1000)
                    {
                        sysAuditResults.InternetUpResult = "N";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "; " + "Upload speed insufficient / Minor issues";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Y";
                    }
                }
                else
                {


                    if (sysAuditResults.cInternetUp[0] < 1)
                    {
                        sysAuditResults.InternetUpResult = "N";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "; " + "Download speed insufficient / Minor issues";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Y";
                    }

                }

                sysAuditResults.cInternetDown = l.ToString().Substring(34, 10);


                if (sysAuditResults.cInternetDown[0] < 3)
                {
                    sysAuditResults.InternetDownResult = "N";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "; " + "Download speed insufficient / Minor issues";
                }
                else
                {
                    sysAuditResults.InternetDownResult = "Y";
                }
            }

            if (l.Contains("CPU"))
            {
                sysAuditResults.cCPU = l.ToString().Substring(26);

                string[] cCPUScores = sysAuditResults.cCPU.Split('=');
                string cCPUScore1 = cCPUScores[1].Replace(" ", "");
                cCPUScore1 = cCPUScores[1].Replace("] - [Processor", "");

                if(!cCPUScore1.Contains("given")) {
                    if (Convert.ToDouble(cCPUScore1) < 4.8)
                    {
                        sysAuditResults.CPUaResult = "N";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "; " + "CPU insufficient - Hardware / System issues";
                    }
                    else
                    {
                        sysAuditResults.CPUaResult = "Y";
                    }
                }
            }



            if (sysAuditResults.OSaResult == "Y" & sysAuditResults.CPUaResult == "Y" & sysAuditResults.RAMaResult == "Y" & sysAuditResults.InternetUpResult == "Y" & sysAuditResults.InternetDownResult == "Y" & sysAuditResults.HDDaResult == "Y")
                sysAuditResults.aResult = "Passed";
            else
            {

                if (sysAuditResults.OSaResult != "Y")
                {
                    sysAuditResults.aResult = sysAuditResults.OSaResult;
                }
                if (sysAuditResults.CPUaResult != "Y")
                {
                    sysAuditResults.aResult = sysAuditResults.CPUaResult;
                }
                if (sysAuditResults.RAMaResult != "Y")
                {
                    sysAuditResults.aResult = sysAuditResults.RAMaResult;
                }
                if (sysAuditResults.InternetUpResult != "Y")
                {
                    sysAuditResults.aResult = sysAuditResults.InternetUpResult;
                }
                if (sysAuditResults.InternetDownResult != "Y")
                {
                    sysAuditResults.aResult = sysAuditResults.InternetDownResult;
                }
                if (sysAuditResults.HDDaResult != "Y")
                {
                    sysAuditResults.aResult = sysAuditResults.HDDaResult;
                }
            }

            // Return the completed audit results
            sysAuditResults.aResultSummary = sysAuditResults.cOS + ", " + sysAuditResults.cCPU + ", " + sysAuditResults.cRAM + ", " + sysAuditResults.cHDD + ", " + sysAuditResults.cInternetDown + ", " + sysAuditResults.cInternetUp;
            return sysAuditResults;
        }

        //Get unseen emails from Sysaudit mailbox
        private static System.Net.Mail.MailMessage[] getMailMessages(S22.Imap.ImapClient imapClient)
        {
            IEnumerable<uint> uids = imapClient.Search(S22.Imap.SearchCondition.Unseen().And(S22.Imap.SearchCondition.From("systemaudit@bit-lever.com")));
            IEnumerable<System.Net.Mail.MailMessage> messages = imapClient.GetMessages(uids);
            return (System.Net.Mail.MailMessage[])messages;
        }

        //Declare struct to hold the SysAudit 
        struct SysAuditResults
        {
            private string date;
            private string name;
            private string email;
            private string host;
            private string hdd;
            private string cpu;
            private string OS;
            private string ram;
            private string internetUp;
            private string internetDown;
            private string results;
            private string resultssummary;

            private string cpuresult;
            private string osresults;
            private string ramresults;
            private string hddresults;
            private string internetupresults;
            private string internetdownresults;

            private string afailedreason;
            public string auditDate
            {
                get
                {
                    return date;
                }
                set
                {

                    date = value;
                }
            }
            public string cName
            {
                get
                {
                    return name;
                }
                set
                {

                    name = value;
                }
            }
            public string cEmail
            {
                get
                {
                    return email;
                }
                set
                {

                    email = value;
                }
            }
            public string cHost
            {
                get
                {
                    return host;
                }
                set
                {

                    host = value;
                }
            }
            public string cHDD
            {
                get
                {
                    return hdd;
                }
                set
                {

                    hdd = value;
                }
            }
            public string cCPU
            {
                get
                {
                    return cpu;
                }
                set
                {

                    cpu = value;
                }
            }
            public string cOS
            {
                get
                {
                    return OS;
                }
                set
                {

                    OS = value;
                }
            }
            public string cRAM
            {
                get
                {
                    return ram;
                }
                set
                {

                    ram = value;
                }
            }
            public string cInternetUp
            {
                get
                {
                    return internetUp;
                }
                set
                {

                    internetUp = value;
                }
            }
            public string cInternetDown
            {
                get
                {
                    return internetDown;
                }
                set
                {

                    internetDown = value;
                }
            }
            public string aResult
            {
                get
                {
                    return results;
                }
                set
                {

                    results = value;
                }
            }
            public string aResultSummary
            {
                get
                {
                    return resultssummary;
                }
                set
                {

                    resultssummary = value;
                }
            }

            public string OSaResult
            {
                get
                {
                    return osresults;
                }
                set
                {

                    osresults = value;
                }
            }
            public string CPUaResult
            {
                get
                {
                    return cpuresult;
                }
                set
                {

                    cpuresult = value;
                }
            }
            public string RAMaResult
            {
                get
                {
                    return ramresults;
                }
                set
                {

                    ramresults = value;
                }
            }
            public string InternetUpResult
            {
                get
                {
                    return internetupresults;
                }
                set
                {

                    internetupresults = value;
                }
            }
            public string InternetDownResult
            {
                get
                {
                    return internetdownresults;
                }
                set
                {

                    internetdownresults = value;
                }
            }
            public string HDDaResult
            {
                get
                {
                    return hddresults;
                }
                set
                {

                    hddresults = value;
                }
            }

            public string aFailedReason
            {
                get
                {
                    return afailedreason;
                }
                set
                {

                    afailedreason = value;
                }
            }

        }
        
        public static void sendMail(string recipient, string attachmentFilename, string cadidateName, string cemail)
        {
            try
            {
                string fileName = @"\\Filesvr4\it\WinAudit\Results_Archive\" + cadidateName.Trim() + " SystemAudit Results.pdf";

                bool ex = File.Exists(fileName);

                if (ex)
                {

                    SmtpClient mySmtpClient = new SmtpClient("secure.emailsrvr.com", 25);
                    mySmtpClient.UseDefaultCredentials = false;
                    System.Net.NetworkCredential basicAuthenticationInfo = new
                    System.Net.NetworkCredential("notify@statesidebpo.com", "W31is+en2016");
                    mySmtpClient.Credentials = basicAuthenticationInfo;

                    // add from,to mailaddresses
                    MailAddress from = new MailAddress("notify@statesidebpo.com");
                    MailAddress to = new MailAddress("brodriguez@statesidebpo.com");
                    MailMessage myMail = new MailMessage(from, to);
                    myMail.IsBodyHtml = true;
                    myMail.Subject = "System audit results";

                   string body = @"<p style =""font-size=16px"">Dear Candidate,<br><br>" + "This email is to inform you of your system audit results. Please see the attachment. If you have any technical questions regarding your results, please reach out to us via email at <a mailto:winaudit@statesidebpo.com>winaudit@statesidebpo.com</a>.</p>";
                   body = body + @"<p style =""font-size=14px""><i>(Please note: If you are unable to view the attachment, you may need to download and install Adobe Acrobat Reader DC or a similar program that allows the viewing of PDF documents)</i>";

                    Attachment inlineStatetSideLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\StatesideLogo.png");
                    Attachment inlineBitLeverLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\Bit-LeverLogo.png");

                    Attachment SysAuditResults = new Attachment(@"\\Filesvr4\it\WinAudit\Results_Archive\" + cadidateName + " SystemAudit Results.pdf");                   
                    myMail.Attachments.Add(SysAuditResults);
                    myMail.Attachments.Add(inlineStatetSideLogo);
                    myMail.Attachments.Add(inlineBitLeverLogo);


                    inlineBitLeverLogo.ContentId = "BitLeverLogo";
                    inlineStatetSideLogo.ContentId =" StatetSideLogo";
                    SysAuditResults.ContentId = "Pdf";

                    inlineBitLeverLogo.ContentDisposition.Inline = true;
                    inlineStatetSideLogo.ContentDisposition.Inline = true;

                    inlineBitLeverLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline;
                    inlineStatetSideLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline;



                    myMail.Body  =  @"<htm><body>" + body + "<br>  <table><tr><td><img src=\"cid:StatetSideLogo\"></td><td><img src=\"cid:BitLeverLogo\"></td></tr></table></body></html>";
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
        public static void sendErrorMail(string recipient, string cadidateName)
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
                MailAddress to = new MailAddress(recipient);
                MailMessage myMail = new MailMessage(from, to);
                myMail.IsBodyHtml = true;
                myMail.Subject = "Processing the system audit for " + cadidateName + " has failed.";
                mySmtpClient.Send(myMail);
            }
            catch (System.Exception ex)
            {
                throw new ApplicationException
                  ("Outlook exception has occured: " + ex.Message);
            }
        }
        private static void sendCompletionNotification(List<SysAuditResults> CandidatesList)
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
                MailAddress to = new MailAddress("helpdesk@statesidebpo.com");
                MailMessage myMail = new MailMessage(from, to);
                myMail.IsBodyHtml = true;
                myMail.Subject = DateTime.Now + " SystemAudit Processing run completed successfully";
                mySmtpClient.Send(myMail);
                
                string bd = "";

                bd = CandidatesList.Count() + " SystemAudit(s) processed.<br/><ol>";
                foreach (SysAuditResults c in CandidatesList)
                {
                    bd = bd + " <li>" + c.cName + " - " + c.aResult + "</li>";
                }
                bd = bd + " </ol>";

                myMail.Body = Regex.Replace(bd, @"[^\u0000-\u007F]", " ");
                mySmtpClient.Send(myMail);

            }
            catch (System.Exception ex)
            {
                throw new ApplicationException
                  ("Outlook exception has occured: " + ex.Message);
            }
        }

        private static void killExcel()
        {
            var process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (var p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }
        class HtmlToText
        {
            // Static data tables
            protected static Dictionary<string, string> _tags;
            protected static HashSet<string> _ignoreTags;

            // Instance variables
            protected TextBuilder _text;
            protected string _html;
            protected int _pos;

            // Static constructor (one time only)
            static HtmlToText()
            {
                _tags = new Dictionary<string, string>();
                _tags.Add("address", "\n");
                _tags.Add("blockquote", "\n");
                _tags.Add("div", "\n");
                _tags.Add("dl", "\n");
                _tags.Add("fieldset", "\n");
                _tags.Add("form", "\n");
                _tags.Add("h1", "\n");
                _tags.Add("/h1", "\n");
                _tags.Add("h2", "\n");
                _tags.Add("/h2", "\n");
                _tags.Add("h3", "\n");
                _tags.Add("/h3", "\n");
                _tags.Add("h4", "\n");
                _tags.Add("/h4", "\n");
                _tags.Add("h5", "\n");
                _tags.Add("/h5", "\n");
                _tags.Add("h6", "\n");
                _tags.Add("/h6", "\n");
                _tags.Add("p", "\n");
                _tags.Add("/p", "\n");
                _tags.Add("table", "\n");
                _tags.Add("/table", "\n");
                _tags.Add("ul", "\n");
                _tags.Add("/ul", "\n");
                _tags.Add("ol", "\n");
                _tags.Add("/ol", "\n");
                _tags.Add("/li", "\n");
                _tags.Add("br", "\n");
                _tags.Add("/td", "\t");
                _tags.Add("/tr", "\n");
                _tags.Add("/pre", "\n");

                _ignoreTags = new HashSet<string>();
                _ignoreTags.Add("script");
                _ignoreTags.Add("noscript");
                _ignoreTags.Add("style");
                _ignoreTags.Add("object");
            }

            /// <summary>
            /// Converts the given HTML to plain text and returns the result.
            /// </summary>
            /// <param name="html">HTML to be converted</param>
            /// <returns>Resulting plain text</returns>
            public string Convert(string html)
            {
                // Initialize state variables
                _text = new TextBuilder();
                _html = html;
                _pos = 0;

                // Process input
                while (!EndOfText)
                {
                    if (Peek() == '<')
                    {
                        // HTML tag
                        bool selfClosing;
                        string tag = ParseTag(out selfClosing);

                        // Handle special tag cases
                        if (tag == "body")
                        {
                            // Discard content before <body>
                            _text.Clear();
                        }
                        else if (tag == "/body")
                        {
                            // Discard content after </body>
                            _pos = _html.Length;
                        }
                        else if (tag == "pre")
                        {
                            // Enter preformatted mode
                            _text.Preformatted = true;
                            EatWhitespaceToNextLine();
                        }
                        else if (tag == "/pre")
                        {
                            // Exit preformatted mode
                            _text.Preformatted = false;
                        }

                        string value;
                        if (_tags.TryGetValue(tag, out value))
                            _text.Write(value);

                        if (_ignoreTags.Contains(tag))
                            EatInnerContent(tag);
                    }
                    else if (Char.IsWhiteSpace(Peek()))
                    {
                        // Whitespace (treat all as space)
                        _text.Write(_text.Preformatted ? Peek() : ' ');
                        MoveAhead();
                    }
                    else
                    {
                        // Other text
                        _text.Write(Peek());
                        MoveAhead();
                    }
                }
                // Return result
                return HttpUtility.HtmlDecode(_text.ToString());
            }

            // Eats all characters that are part of the current tag
            // and returns information about that tag
            protected string ParseTag(out bool selfClosing)
            {
                string tag = String.Empty;
                selfClosing = false;

                if (Peek() == '<')
                {
                    MoveAhead();

                    // Parse tag name
                    EatWhitespace();
                    int start = _pos;
                    if (Peek() == '/')
                        MoveAhead();
                    while (!EndOfText && !Char.IsWhiteSpace(Peek()) &&
                        Peek() != '/' && Peek() != '>')
                        MoveAhead();
                    tag = _html.Substring(start, _pos - start).ToLower();

                    // Parse rest of tag
                    while (!EndOfText && Peek() != '>')
                    {
                        if (Peek() == '"' || Peek() == '\'')
                            EatQuotedValue();
                        else
                        {
                            if (Peek() == '/')
                                selfClosing = true;
                            MoveAhead();
                        }
                    }
                    MoveAhead();
                }
                return tag;
            }

            // Consumes inner content from the current tag
            protected void EatInnerContent(string tag)
            {
                string endTag = "/" + tag;

                while (!EndOfText)
                {
                    if (Peek() == '<')
                    {
                        // Consume a tag
                        bool selfClosing;
                        if (ParseTag(out selfClosing) == endTag)
                            return;
                        // Use recursion to consume nested tags
                        if (!selfClosing && !tag.StartsWith("/"))
                            EatInnerContent(tag);
                    }
                    else MoveAhead();
                }
            }

            // Returns true if the current position is at the end of
            // the string
            protected bool EndOfText
            {
                get { return (_pos >= _html.Length); }
            }

            // Safely returns the character at the current position
            protected char Peek()
            {
                return (_pos < _html.Length) ? _html[_pos] : (char)0;
            }

            // Safely advances to current position to the next character
            protected void MoveAhead()
            {
                _pos = Math.Min(_pos + 1, _html.Length);
            }

            // Moves the current position to the next non-whitespace
            // character.
            protected void EatWhitespace()
            {
                while (Char.IsWhiteSpace(Peek()))
                    MoveAhead();
            }

            // Moves the current position to the next non-whitespace
            // character or the start of the next line, whichever
            // comes first
            protected void EatWhitespaceToNextLine()
            {
                while (Char.IsWhiteSpace(Peek()))
                {
                    char c = Peek();
                    MoveAhead();
                    if (c == '\n')
                        break;
                }
            }

            // Moves the current position past a quoted value
            protected void EatQuotedValue()
            {
                char c = Peek();
                if (c == '"' || c == '\'')
                {
                    // Opening quote
                    MoveAhead();
                    // Find end of value
                    int start = _pos;
                    _pos = _html.IndexOfAny(new char[] { c, '\r', '\n' }, _pos);
                    if (_pos < 0)
                        _pos = _html.Length;
                    else
                        MoveAhead();    // Closing quote
                }
            }

            /// <summary>
            /// A StringBuilder class that helps eliminate excess whitespace.
            /// </summary>
            protected class TextBuilder
            {
                private StringBuilder _text;
                private StringBuilder _currLine;
                private int _emptyLines;
                private bool _preformatted;

                // Construction
                public TextBuilder()
                {
                    _text = new StringBuilder();
                    _currLine = new StringBuilder();
                    _emptyLines = 0;
                    _preformatted = false;
                }

                public bool Preformatted
                {
                    get
                    {
                        return _preformatted;
                    }
                    set
                    {
                        if (value)
                        {
                            // Clear line buffer if changing to
                            // preformatted mode
                            if (_currLine.Length > 0)
                                FlushCurrLine();
                            _emptyLines = 0;
                        }
                        _preformatted = value;
                    }
                }

                public void Clear()
                {
                    _text.Length = 0;
                    _currLine.Length = 0;
                    _emptyLines = 0;
                }

                public void Write(string s)
                {
                    foreach (char c in s)
                        Write(c);
                }

                public void Write(char c)
                {
                    if (_preformatted)
                    {
                        // Write preformatted character
                        _text.Append(c);
                    }
                    else
                    {
                        if (c == '\r')
                        {
                            // Ignore carriage returns. We'll process
                            // '\n' if it comes next
                        }
                        else if (c == '\n')
                        {
                            // Flush current line
                            FlushCurrLine();
                        }
                        else if (Char.IsWhiteSpace(c))
                        {
                            // Write single space character
                            int len = _currLine.Length;
                            if (len == 0 || !Char.IsWhiteSpace(_currLine[len - 1]))
                                _currLine.Append(' ');
                        }
                        else
                        {
                            // Add character to current line
                            _currLine.Append(c);
                        }
                    }
                }

                protected void FlushCurrLine()
                {
                    // Get current line
                    string line = _currLine.ToString().Trim();

                    // Determine if line contains non-space characters
                    string tmp = line.Replace("&nbsp;", String.Empty);
                    if (tmp.Length == 0)
                    {
                        // An empty line
                        _emptyLines++;
                        if (_emptyLines < 2 && _text.Length > 0)
                            _text.AppendLine(line);
                    }
                    else
                    {
                        // A non-empty line
                        _emptyLines = 0;
                        _text.AppendLine(line);
                    }

                    // Reset current line
                    _currLine.Length = 0;
                }

                public override string ToString()
                {
                    if (_currLine.Length > 0)
                        FlushCurrLine();
                    return _text.ToString();
                }
            }
        }
        class Excell
        {
            public void openExcelFile()
            {
                Excel.Application oXL = new Excel.Application();

                Excel.Workbook oWB = oXL.Workbooks.Open(@"C:\Winaudit\", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                //read Excel sheets 
                foreach (Excel.Worksheet ws in oWB.Sheets)
                {
                    Console.WriteLine(ws.Name);
                }

                //save as separate copy 
                //oWB.SaveAs(Application.StartupPath + "\\PROJEKTSTATUS_GESAMT_neues_Layout_neu.xlsm", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                oWB.Close(true, Missing.Value, Missing.Value);
            }

            public void writeExcelFile()
            {
                Excel.Application oXL = new Excel.Application();
                Excel.Workbook oWB = oXL.Workbooks.Open(@"C:\Winaudit\WinAuditPro_4.0.xlsm", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                Excel.Worksheet oWS = oWB.Worksheets[1] as Excel.Worksheet;

                for (int i = 1; i < 10; i++)
                {
                    oWS.Cells[i, 1] = "Cell " + i.ToString();
                }

            }
        }

    }
}
