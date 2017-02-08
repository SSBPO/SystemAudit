using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using AE.Net.Mail;
using System.Web;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
//using S22.Imap;
using System.Reflection;
using System.Net.Mail;
using Microsoft.Vbe.Interop;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Net;

namespace StatesideBpo
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
            Console.WriteLine("#######################      SysAuditPro  ver.2    #############################");
            Console.WriteLine("");

            // Declare app instance / list to hold all sysaudits 
            SystemAudit xxx_SysAudit = new SystemAudit();
            List<SysAuditResults> CandidatesList = new List<SysAuditResults>();
            List<string> CandidatesValueList = new List<string>();

            using (ImapClient imapClient = new ImapClient("secure.emailsrvr.com", "systemaudit@statesidebpo.com", "g4d5fg4df!!2 ew", AuthMethods.Login, 993, true))
            {
                killExcel();
                int ProcessedEmails = 0;
                HtmlToText stripHtml = new HtmlToText();

                Lazy<AE.Net.Mail.MailMessage>[] msgs = getMailMessages(imapClient);   //Need to improve this
                string s = "Checking SystemAudits mailbox...";
                Console.WriteLine(s);
                Console.WriteLine("");

                if (msgs.Count() != 0)
                {
                    if (msgs.Count() == 1)
                    {
                       Console.WriteLine("There is " + msgs.Count() + " SystemAudit to process.");
                    }
                    else
                    {
                        Console.WriteLine("There are " + msgs.Count() + " SystemAudits to process.");
                    }
                    Console.WriteLine("");
                    Console.WriteLine("Starting...");
                    Console.WriteLine("");

                    //Foreach unseen email found in the mailbox
                    foreach (Lazy<AE.Net.Mail.MailMessage> msg in msgs)
                    {
                        //Declare sysaudit results object
                        SysAuditResults SSBPOsysAuditResults = new SysAuditResults();
                       

                        // Flag each email as seen
                        imapClient.AddFlags(Flags.Seen, msg.Value);   

                        if ((msg.Value.Body != "") && msg.Value != null)
                        {
                            ProcessedEmails = ProcessedEmails + 1;

                            string[] lines = stripHtml.Convert(msg.Value.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                            int i = 2;
                            List<String> line = lines.ToList();

                            SysAuditXLWApp = new Excel.Application();
                            SysAuditWWorkBook = SysAuditXLWApp.Workbooks.Open(@"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xltm"); // Open the SysAudit Excel template
                            SysAuditWWorkSheet = SysAuditWWorkBook.Worksheets[1] as Excel.Worksheet; // Set sheet 1 as the active sheet in Excel template

                            string s1 = "Processing " + ProcessedEmails + " of " + msgs.Count() + ".";
                            Console.WriteLine(s1);

                            foreach (string l in line.ToList()) //Foreach of the lines in the email
                            {
                                if (l.ToString() != "" & !l.ToString().Contains("Simplified Audit Results") & !l.ToString().Contains("www") & !l.Contains("Avast") & !l.Contains("antivirus")) //If the line is not empty or has unwanted text
                                {
                                    SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, l); // Set the results object varialble to the value on the specified in the line 
                                    SysAuditWWorkSheet.Cells[i, 2] = l.ToString().TrimStart(); //Write the line value to the active sheet in Excel template
                                    CandidatesValueList.Add(l.ToString().TrimStart());
                                    i++;
                                }
                            }

                            if (!SSBPOsysAuditResults.cCPU.Contains("given"))
                            {
                                SysAuditXLWApp.Run("Sheet2.SaveAsC");  //Run SaveAsC macro on the Excel template to export results to pdf                                 // 
                            }

                            if (SSBPOsysAuditResults.cCPU.Contains("given"))
                            {
                                    SSBPOsysAuditResults.aFailedReason = "**Needs manual processing**";                                
                            }

                            object misValue = System.Reflection.Missing.Value; //Get misssing.vlaue variable
                            SysAuditWWorkBook.Close(false, misValue, misValue);
                            SysAuditWWorkSheet = null;
                            SysAuditWWorkBook = null;
                            SysAuditXLWApp.Quit();                           
                        }

                        CandidatesList.Add(SSBPOsysAuditResults);
                    }
                }
                if (CandidatesList.Count() > 0)
                {
                    createBitLeverImport(CandidatesList);
                    sendCompletionNotification(CandidatesList);                  
                }

            }/// Using imap
            Console.WriteLine("");
            Console.WriteLine("All WinAudits have been processed.");
            Thread.Sleep(5000);
        }//Main
        
        private static void createBitLeverImport(List<SysAuditResults> CandidatesList)
        {
            string file2Import = string.Format(@"\\filesvr4\IT\WinAudit\4BitLeverImport\BitLeverImport{0:yyyy-MM-dd_hh-mm-ss-tt}" + " Results.xls", DateTime.Now);
            object misValue = System.Reflection.Missing.Value;
            SysAuditXLWApp2 = new Excel.Application();
            SysAuditXLWApp2.DisplayAlerts = false;
            SysAuditWWorkBook2 = SysAuditXLWApp2.Workbooks.Add(misValue);
            SysAuditWWorkSheet2 = SysAuditWWorkBook2.Worksheets[1] as Excel.Worksheet;

            using (var stream = File.CreateText(file2Import))
            {
                Microsoft.Office.Interop.Excel.Range range = SysAuditWWorkSheet2.UsedRange;
                int colCount = range.Columns.Count;
                int rowCount = range.Rows.Count;
                int index = 2;

                SysAuditWWorkSheet2.Cells[1, 1] = "Date Processed";
                SysAuditWWorkSheet2.Cells[1, 2] = "Candidate Name";
                SysAuditWWorkSheet2.Cells[1, 3] = "Candidate Email";
                SysAuditWWorkSheet2.Cells[1, 4] = "Notes";
                SysAuditWWorkSheet2.Cells[1, 5] = "SysAudit Status";
                SysAuditWWorkSheet2.Cells[1, 6] = "Processed By";
                SysAuditWWorkSheet2.Cells[1, 7] = "Results Sent?";
                SysAuditWWorkSheet2.Cells[1, 8] = "Fail Reason";

                foreach (SysAuditResults r in CandidatesList)
                {
                        SysAuditWWorkSheet2.Cells[index, 1] = r.auditDate;
                        SysAuditWWorkSheet2.Cells[index, 2] = r.cName;
                        SysAuditWWorkSheet2.Cells[index, 3] = r.cEmail;
                        SysAuditWWorkSheet2.Cells[index, 4] = r.aResultSummary;
                        SysAuditWWorkSheet2.Cells[index, 5] = r.aResult.Replace("ed", "").Replace(".","").Replace('-', ' ');
                        SysAuditWWorkSheet2.Cells[index, 6] = Environment.UserName;
                        SysAuditWWorkSheet2.Cells[index, 7] = "Yes";
                        if (r.aFailedReason != null)
                            SysAuditWWorkSheet2.Cells[index, 8] = r.aFailedReason;
                        index = index + 1;

                }

            }

            SysAuditWWorkBook2.Close(SaveChanges: true, Filename: file2Import);
            SysAuditWWorkSheet2 = null;
            SysAuditWWorkBook2 = null;
            SysAuditXLWApp2.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }        
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
                    sysAuditResults.HDDaResult = "Fail";
                    sysAuditResults.aResult = "Fail";
                    sysAuditResults.aFailedReason =  "Hard drive insufficient";
                }
                else
                {
                    sysAuditResults.HDDaResult = "Pass";
                }
            }

            if (l.Contains("Operating"))
            {
                sysAuditResults.cOS = l.ToString().Substring(18);
                string cOSs = sysAuditResults.cOS.Substring(0, 10);//.Split('s');
                string[] cOS1s = cOSs.Split(' ');
                string cOS2 = cOS1s[1];
                //"8.1 Build 9200"

                if (cOS2.Trim().Length > 1)
                {
                    cOS1s = cOS2.Split('.');
                    cOS2 = cOS1s[0];
                }


                if (Convert.ToInt64(cOS2.Trim()) < 7)
                {
                    sysAuditResults.OSaResult = "Fail";
                    sysAuditResults.aResult = "Fail";
                    sysAuditResults.aFailedReason = "OS insufficient";
                }
                else
                {
                    sysAuditResults.OSaResult = "Pass";
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
                    sysAuditResults.RAMaResult = "Fail";
                    sysAuditResults.aResult = "Fail";
                    sysAuditResults.aFailedReason = "RAM insufficient ";
                }
                else
                {
                    sysAuditResults.RAMaResult = "Pass";
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
                        sysAuditResults.InternetUpResult = "Fail";
                        sysAuditResults.aResult = "Fail";
                        sysAuditResults.aFailedReason = "Upload speed insufficient";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Pass";
                    }
                }
                else
                {
                    if (sysAuditResults.cInternetUp[0] < 1)
                    {
                        sysAuditResults.InternetUpResult = "Fail";
                        sysAuditResults.aResult = "Fail";
                        sysAuditResults.aFailedReason = "Download speed insufficient";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Pass";
                    }

                }

                sysAuditResults.cInternetDown = l.ToString().Substring(34, 10);


                if (sysAuditResults.cInternetDown[0] < 3)
                {
                    sysAuditResults.InternetDownResult = "Fail";
                    sysAuditResults.aResult = "Fail";
                    sysAuditResults.aFailedReason = "Download speed insufficient ";
                }
                else
                {
                    sysAuditResults.InternetDownResult = "Pass";
                }
            }

            if (l.Contains("CPU"))
            {
                sysAuditResults.cCPU = l.ToString().Substring(26);

                string[] cCPUScores = sysAuditResults.cCPU.Split('=');
                string cCPUScore1 = cCPUScores[1].Replace(" ", "");
                cCPUScore1 = cCPUScores[1].Replace("] - [Processor", "").Trim();
                if (!cCPUScore1.Contains("given"))
                {
                    if (Convert.ToDouble(cCPUScore1) > 4.8)
                    {
                        sysAuditResults.CPUaResult = "Pass";
                    }
                    else
                    {
                        sysAuditResults.CPUaResult = "Fail";
                        sysAuditResults.aFailedReason = "CPU insufficient";
                        sysAuditResults.aResult = "Fail";
                    }
                }
                else
                {
                    sysAuditResults.CPUaResult = "Fail";
                    sysAuditResults.aFailedReason = "**Needs manual processing**";
                    sysAuditResults.needsManualProcessing = true;
                    sysAuditResults.aResult = "Fail";
                }
            }


            sysAuditResults = ProcessAudits(sysAuditResults);

            sysAuditResults.aResultSummary = sysAuditResults.cOS + ", " + sysAuditResults.cCPU + ", " + sysAuditResults.cRAM + ", " + sysAuditResults.cHDD + ", " + "[Download Speed: " + sysAuditResults.cInternetDown + ", " + "[Upload Speed: " + sysAuditResults.cInternetUp + "]";



            return sysAuditResults;
        }
        private static SysAuditResults ProcessAudits(SysAuditResults sysAuditResults)
        {

            if (sysAuditResults.needsManualProcessing == false)
            {

                if (sysAuditResults.OSaResult == "Pass" && sysAuditResults.CPUaResult == "Pass" && sysAuditResults.RAMaResult == "Pass" && sysAuditResults.InternetUpResult == "Pass" && sysAuditResults.InternetDownResult == "Pass" && sysAuditResults.HDDaResult == "Pass")
                {
                    sysAuditResults.aResult = "Pass";
                }

                else
                {

                    if (sysAuditResults.OSaResult != "Pass")
                    {
                        sysAuditResults.aResult = sysAuditResults.OSaResult;

                    }
                    if (sysAuditResults.CPUaResult != "Pass")
                    {
                        sysAuditResults.aResult = sysAuditResults.CPUaResult;
                    }

                    if (sysAuditResults.RAMaResult != "Pass")
                    {
                        sysAuditResults.aResult = sysAuditResults.RAMaResult;
                    }
                    if (sysAuditResults.InternetUpResult != "Pass")
                    {
                        sysAuditResults.aResult = sysAuditResults.InternetUpResult;
                    }
                    if (sysAuditResults.InternetDownResult != "Pass")
                    {
                        sysAuditResults.aResult = sysAuditResults.InternetDownResult;
                    }
                    if (sysAuditResults.HDDaResult != "Pass")
                    {
                        sysAuditResults.aResult = sysAuditResults.HDDaResult;
                    }
                }
            }

            return sysAuditResults;
        }
        private static Lazy<AE.Net.Mail.MailMessage>[] getMailMessages(ImapClient imapClient)
        {
            imapClient.SelectMailbox("INBOX");
            Regex regex = new Regex(@":");
            // Lazy<AE.Net.Mail.MailMessage>[] messages = imapClient.SearchMessages(SearchCondition.From("systemaudit@bit-lever.com"), false);
            Lazy<AE.Net.Mail.MailMessage>[] messages = imapClient.SearchMessages(SearchCondition.From("systemaudit@bit-lever.com").And(SearchCondition.Unseen()));

            return messages;
        }

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
            private bool needsmanualprocessing;

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

            public bool needsManualProcessing
            {
                get
                {
                    return needsmanualprocessing;
                }
                set
                {

                    needsmanualprocessing = value;
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
        public static Outlook.Account getAccountForEmailAddress(Outlook.Application application, string smtpAddress)
        {
            // Loop over the Accounts collection of the current Outlook session. 
            Outlook.Accounts accounts = application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                // When the e-mail address matches, return the account. 
                if (account.SmtpAddress == smtpAddress)
                {
                    return account;
                }
            }
            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", smtpAddress));
        }
        public static void sendMail(string recipient, string attachmentFilename, string cadidateName)
        {
            if (File.Exists(attachmentFilename)) {
                    try
                    {
                        Outlook.Application otApp = new Outlook.Application();// create outlook object
                        Outlook.NameSpace ns = otApp.Session;

                        Outlook.Folder folder = otApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts) as Outlook.Folder;
                        Outlook.MailItem otMsg = otApp.CreateItemFromTemplate(@"\\filesvr4\IT\WinAudit\SysAudit App\System audit results.oft", folder) as Outlook.MailItem;///);Outlook.MailItem)otApp.CreateItem(Outlook.OlItemType.olMailItem); // Create mail object

                        otMsg.SendUsingAccount = getAccountForEmailAddress(otApp, "systemaudit@statesidebpo.com");

                        Outlook.Inspector oInspector = otMsg.GetInspector;
                        Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add(recipient);


                        otMsg.Recipients.ResolveAll();// validate recipient address

                        otMsg.Subject = "SSBPO System audit results";
                        String sSource = attachmentFilename;
                        String sDisplayName = cadidateName + " SystemAudit Results.pdf";

                        int iPos = (int)otMsg.Body.Length + 1;
                        int iAttType = (int)Outlook.OlAttachmentType.olByValue;
                        Outlook.Attachment oAttach = otMsg.Attachments.Add(sSource, iAttType, iPos, sDisplayName); // add attachment
                        otMsg.Save();
                        otMsg.Send(); // Send Mail
                        otRecip = null;
                        // otAttach = null;
                        otMsg = null;
                        otApp = null;

                    }
                    catch (System.Exception ex)
                    {
                        throw new ApplicationException
                          ("Outlook exception has occured: " + ex.Message);
                    }
            }
        }
       
        private static void sendCompletionNotification(List<SysAuditResults> CandidatesList)
        {
            try
            {
                Outlook.Application otApp = new Outlook.Application();// create outlook object
                Outlook.NameSpace ns = otApp.Session;
                Outlook.MailItem otMsg = otApp.CreateItem(Outlook.OlItemType.olMailItem); // Create mail object
                Outlook.Inspector oInspector = otMsg.GetInspector;
                otMsg.SendUsingAccount = getAccountForEmailAddress(otApp, "systemaudit@statesidebpo.com");
                Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add("helpdesk@statesidebpo.com");
                Outlook.Recipient recipBcc = otMsg.Recipients.Add("recruiters@statesidebpo.com");
                recipBcc.Type = (int)Outlook.OlMailRecipientType.olBCC;


                otRecip.Resolve();// validate recipient address
                string bd = "";
                if (CandidatesList.Count() == 1)
                {
                    bd = "<h2> " + CandidatesList.Count() + " Audit was processed</h2>";
                }
                else
                {
                    bd = "<h2> " + CandidatesList.Count() + " Audits were processed</h2>";
                }

                string attachmentFilename = "";
                bd = bd + "<table border = " + "1" + " cellpadding = " + "6" + " cellspacing = " + "5" + "><tbody>";
             

                foreach (SysAuditResults c in CandidatesList)
                {
                    attachmentFilename = @"\\filesvr4\IT\WinAudit\Results_Archive\" + c.cName + " SystemAudit Results.pdf";
                    string reason = c.aFailedReason;
                    if (c.aFailedReason != null)
                    {
                        reason = " - " + c.aFailedReason;
                    }

                    if (c.needsManualProcessing) {        

                        bd = bd + "<tr><td  width=" + "'45%'" + ">" + c.cName + "</td><td width=" + "'55%'" + ">" + c.cEmail + "</td><td width=" + "'55%'" + ">" + c.aResult + reason + "</td></tr>";
                        sendMail(c.cEmail, attachmentFilename, c.cName);
                    }
                    else
                    {
                        bd = bd + "<tr><td  width=" + "'45%'" + ">" + c.cName + "</td><td width=" + "'55%'" + ">" + c.cEmail + "</td><td width=" + "'55%'" + ">"  + reason + "</td></tr>";
                        sendMail(c.cEmail, attachmentFilename, c.cName);
                    }
                    
                }

                bd = bd + "</tbody></table>";

                otMsg.HTMLBody = Regex.Replace(bd, @"[^\u0000-\u007F]", " ");             
                otMsg.Subject = DateTime.Now + " SystemAudit Processing run completed successfully";                
                
                otMsg.Send();
                otRecip = null;
                otMsg = null;
                otApp = null;
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

