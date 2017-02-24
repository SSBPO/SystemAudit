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
        public static bool isTESTING { get; private set; }

        public int ProcessedEmails = 0;
        public int ProcessedManualEmails = 0;



        static void Main(string[] args)
        {
            string excelLicense = "EQU2-1K6F-UZPP-4MOX";
            SpreadsheetInfo.SetLicense(excelLicense);
            isTESTING = true;

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("################################################################################");
            Console.WriteLine("#####################         WinAuditPro ver.3       ##########################");
            Console.WriteLine("################################################################################");
            Console.WriteLine();

            // Declare app instance / list to hold all sysaudits 
            SystemAudit xxx_SysAudit = new SystemAudit();
            List<SysAuditResults> CandidatesList = new List<SysAuditResults>();


            using (S22.Imap.ImapClient imapClient = new S22.Imap.ImapClient("secure.emailsrvr.com", 993, "systemaudit@statesidebpo.com", "g4d5fg4df!!2 ew", AuthMethod.Login, true))
            {
                string s = "Checking SystemAudits mailbox...";
                Console.SetCursorPosition((Console.WindowWidth - s.Length) / 2, Console.CursorTop);
                Console.WriteLine(s);
                Console.WriteLine();
               

                // Check mailbox and get any messages not seen and sent by systemaudit@bit-lever.com
                IEnumerable<uint> uids = imapClient.Search(S22.Imap.SearchCondition.Unseen().And(S22.Imap.SearchCondition.From("systemaudit@bit-lever.com")));
                IEnumerable<System.Net.Mail.MailMessage> messages = imapClient.GetMessages(uids);

                IEnumerable<uint> manualuids = imapClient.Search(S22.Imap.SearchCondition.Unseen().And(S22.Imap.SearchCondition.From("no-reply@bit-lever.com")));
                IEnumerable<System.Net.Mail.MailMessage> manualmessages = imapClient.GetMessages(manualuids);
                
                xxx_SysAudit.ProcessedEmails = processNormalAudits(CandidatesList, messages);
                xxx_SysAudit.ProcessedManualEmails = proccessManualAudits(CandidatesList, manualmessages);
            }

            Console.WriteLine();          
            string s5 = "There are not more emails to processed. Good bye!";
            Console.SetCursorPosition((Console.WindowWidth - s5.Length) / 2, Console.CursorTop);
            Console.WriteLine(s5);
            Thread.Sleep(5000);

        }

 
        private static int processNormalAudits(List<SysAuditResults> CandidatesList, IEnumerable<MailMessage> messages)
        {
            int ProcessedEmails = 0;

            if (messages.Count() != 0)
            {
                string welcomeMsg = "";

                // If there are new unseen messages 
                if (messages.Count() == 1)
                {
                    welcomeMsg = "There is " + messages.Count() + " SystemAudit to process.";
                }
                if (messages.Count() >= 1)
                {
                    welcomeMsg = "There are " + messages.Count() + " SystemAudits to process.";
                }
               
                Console.WriteLine();             

                string s4 = "Starting...";
                Console.SetCursorPosition((Console.WindowWidth - s4.Length) / 2, Console.CursorTop);
                Console.WriteLine(s4);
             

                // For each message

                foreach (System.Net.Mail.MailMessage msg in messages)
                {
                    ProcessedEmails = ProcessedEmails + 1;
                    string s5 = "Processing {0} of {1}.";
                    Console.SetCursorPosition((Console.WindowWidth - s5.Length) / 2, Console.CursorTop);
                    Console.Write(s5, ProcessedEmails, messages.Count());

                   // Console.WriteLine("");

                    //Load the WinAudit workbook
                    string fileName = @"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xlsm";
                    var ef = ExcelFile.Load(fileName);

                    //Instanciate SystemAudit Results object
                    SysAuditResults SSBPOsysAuditResults = new SysAuditResults();
                    HtmlToText stripHtml = new HtmlToText();

                    // If message body is not empty 
                    if (msg.Body != "")
                    {

                        //Strip Html from emal and split into 14 lines  
                        string[] line = stripHtml.Convert(msg.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                        //Foreach line in the email
                        foreach (string l in line.ToList())
                        {
                            if (l.ToString() != "" & !l.ToString().Contains("Simplified Audit Results") & !l.ToString().Contains("www") & !l.Contains("Avast") & !l.Contains("antivirus")) //If the line is not empty or has unwanted text
                            {
                                //Add the values to the the SystemAudit results object
                                SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, l);
                            }
                        }
                    }

                    SSBPOsysAuditResults = getAuditStatus(SSBPOsysAuditResults);
                    generatePdf(ref ef, ref SSBPOsysAuditResults);                    
                    CandidatesList.Add(SSBPOsysAuditResults);
          
                }

                createBitLeverImport(CandidatesList);                
                sendCompletionNotification(CandidatesList);


            }
            else
            {
               
                //string s4 = "There is nothing to process. Good bye!";
                //Console.SetCursorPosition((Console.WindowWidth - s4.Length) / 2, Console.CursorTop);
                //Console.WriteLine(s4);
            }

            return ProcessedEmails;
        }

        private static int proccessManualAudits(List<SysAuditResults> CandidatesList, IEnumerable<MailMessage> Manualmsgs)
        {

            int ProcesseManualdEmails = 0;

            if (Manualmsgs.Count() != 0)
            {
                if (Manualmsgs.Count() == 1)
                {

                    string s9 = "There is " + Manualmsgs.Count() + " Manual SystemAudit to process.";
                    Console.SetCursorPosition((Console.WindowWidth - s9.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s9);
                    Console.WriteLine();
                }
                else
                {
                    Console.WriteLine();
                    string s8 = "There are " + Manualmsgs.Count() + " Manual SystemAudits to process.";
                    Console.SetCursorPosition((Console.WindowWidth - s8.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s8);
                    Console.WriteLine();
                }

                string sx = "Starting...";
                Console.SetCursorPosition((Console.WindowWidth - sx.Length) / 2, Console.CursorTop);
                Console.WriteLine(sx);
                Console.WriteLine();


                HtmlToText stripHtml = new HtmlToText();

                //Foreach unseen email found in the mailbox
                foreach (System.Net.Mail.MailMessage msg in Manualmsgs)
                {
                    if (msg.Body != "")
                    {
                        SysAuditResults SSBPOsysAuditResults = new SysAuditResults();
                        string[] lines = stripHtml.Convert(msg.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                        List<String> line = lines.ToList();
                        
                        //Load the WinAudit workbook
                        string fileName = @"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xlsm";
                        var ef = ExcelFile.Load(fileName);
                        ExcelWorksheet worksheet = ef.Worksheets.ActiveWorksheet;

                        string s1 = "Processing {0} of {1}.";
                        Console.SetCursorPosition((Console.WindowWidth - s1.Length) / 2, Console.CursorTop);
                        Console.SetCursorPosition(0, 0);
                        Console.WriteLine("");
                        Console.Write(s1, ProcesseManualdEmails, Manualmsgs.Count());
                        Console.WriteLine("");

                        worksheet.Cells["B5"].Value = "Computer Name: ";
                        worksheet.Cells["B6"].Value = "Current User:";
                        string manualUpSpeed = "";
                        string manualDownSpeed = "";

                        foreach (string l in line.ToList()) //Foreach of the lines in the email
                        {
                            if (l.ToString().Contains("Date Processed"))
                            {
                               
                                string ln = "Audit Date Time: " + l.ToString().TrimStart().Replace("Date Processed", "");
                                worksheet.Cells["B2"].Value = ln;
                                SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                            }

                            if (l.ToString().Contains("Candidate Name"))
                            {
                                string ln = "Full Name: " + l.ToString().TrimStart().Replace("Candidate Name", "");
                                worksheet.Cells["B3"].Value = ln;                                
                                SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                            }

                            if (l.ToString().Contains("Candidate Email"))
                            {
                                string ln = "Email Address: " + l.ToString().TrimStart().Replace("Candidate Email", "");
                                worksheet.Cells["B4"].Value = ln;
                                SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                            }

                            if (l.ToString().Contains("Notes"))
                            {
                                string[] values = l.Split('-');
                                
                                foreach (string t in values)
                                {
                                    string v = t.Replace("Notes", "");

                                    if (v.Contains("Windows"))
                                    {
                                        string ln = "Operating System: " + v.ToString().TrimStart();
                                        worksheet.Cells["B7"].Value = ln;
                                        SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                                    }

                                    if (v.Contains("CPU"))
                                    {
                                        string ln = "CPU (Processor) Results: " + v.ToString().TrimStart();
                                        worksheet.Cells["B8"].Value = ln;
                                        SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                                    }

                                    if (v.Contains("Total RAM"))
                                    {
                                        string ln = "RAM (Memory) Results: " + v.ToString().TrimStart();
                                        worksheet.Cells["B9"].Value = ln;
                                        SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                                    }

                                    if (v.Contains("Total Space"))
                                    {
                                        string ln = "Disk (Hard Drive) Results:  " + v.ToString().TrimStart();
                                        worksheet.Cells["B10"].Value = ln;
                                        SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                                    }

                                    if (v.Contains("Download"))
                                    {
                                        manualDownSpeed = v.ToString().TrimStart() + "] - ";
                                    }

                                    if (v.Contains("Upload"))
                                    {
                                        manualUpSpeed = "[" + v.ToString().TrimStart() + "]";
                                        string ln = "Network Results: " + manualDownSpeed + manualUpSpeed;
                                        worksheet.Cells["B11"].Value = ln;
                                        SSBPOsysAuditResults = getmManualResultsObject(SSBPOsysAuditResults, ln);
                                    }
                                }
                            }
                        }

                        SSBPOsysAuditResults = getAuditStatus(SSBPOsysAuditResults);
                        generatePdf(ref ef, ref SSBPOsysAuditResults);
                        CandidatesList.Add(SSBPOsysAuditResults);
                    }


                }
            }

            return ProcesseManualdEmails;
        }
        private static void generatePdf(ref ExcelFile ef, ref SysAuditResults SSBPOsysAuditResults)
        {
            //Fill out the candidates results worksheet and save as pdf to send out
            
            ExcelWorksheet worksheet = ef.Worksheets.ActiveWorksheet;
            worksheet.Cells["C5"].Value = SSBPOsysAuditResults.cName;
            worksheet.Cells["C6"].Value = SSBPOsysAuditResults.auditDate;
            worksheet.Cells["D6"].Value = SSBPOsysAuditResults.aResult;

            if (SSBPOsysAuditResults.aResult == "Pass")
            {
                worksheet.Cells["D6"].Style.Font.Color = System.Drawing.Color.Green;
            }
            else
            {
                worksheet.Cells["D6"].Style.Font.Color = System.Drawing.Color.Red;
            }
            worksheet.Cells["C13"].Value = SSBPOsysAuditResults.cOS.Substring(0, 10).Replace(".", "").Trim();
            worksheet.Cells["C14"].Value = SSBPOsysAuditResults.cHDD.Substring(64).Replace("]", "").Trim();

            if (SSBPOsysAuditResults.cRAM.Length > 35)
            {
                worksheet.Cells["C15"].Value = SSBPOsysAuditResults.cRAM.Substring(33, 4).Trim().Replace("]", " ");
            }else
            {
                worksheet.Cells["C15"].Value = SSBPOsysAuditResults.cRAM.Substring(30, 4).Trim().Replace("]", " ");
            }

            worksheet.Cells["C16"].Value = SSBPOsysAuditResults.cCPU.Substring(11, 4).Trim();
            worksheet.Cells["C18"].Value = SSBPOsysAuditResults.cInternetUp.Trim() + " Mbps";
            worksheet.Cells["C19"].Value = SSBPOsysAuditResults.cInternetDown.Trim() + " Mbps";
            if (Convert.ToInt32(SSBPOsysAuditResults.cInternetUp) > 99)
            {
                worksheet.Cells["C18"].Value = SSBPOsysAuditResults.cInternetUp.Trim() + " Kbps";
            }
            if (Convert.ToInt32(SSBPOsysAuditResults.cInternetDown) > 99)
            {
                worksheet.Cells["C19"].Value = SSBPOsysAuditResults.cInternetDown.Trim() + " Kbps";
            }

            worksheet.Cells["D13"].Value = SSBPOsysAuditResults.OSaResult;
            worksheet.Cells["D14"].Value = SSBPOsysAuditResults.HDDaResult;
            worksheet.Cells["D15"].Value = SSBPOsysAuditResults.RAMaResult;
            worksheet.Cells["D16"].Value = SSBPOsysAuditResults.CPUaResult;
            worksheet.Cells["D18"].Value = SSBPOsysAuditResults.InternetUpResult;
            worksheet.Cells["D19"].Value = SSBPOsysAuditResults.InternetDownResult;

            if (SSBPOsysAuditResults.OSaResult == "Pass")
            {
                worksheet.Cells["D13"].Value = "Y";
                worksheet.Cells["D13"].Style.Font.Color = System.Drawing.Color.Green;

            }
            else
            {
                worksheet.Cells["D13"].Value = "N";
                worksheet.Cells["D13"].Style.Font.Color = System.Drawing.Color.Red;
            }
            if (SSBPOsysAuditResults.HDDaResult == "Pass")
            {
                worksheet.Cells["D14"].Value = "Y";
                worksheet.Cells["D14"].Style.Font.Color = System.Drawing.Color.Green;
            }
            else
            {
                worksheet.Cells["D14"].Value = "N";
                worksheet.Cells["D14"].Style.Font.Color = System.Drawing.Color.Red;
            }
            if (SSBPOsysAuditResults.RAMaResult == "Pass")
            {
                worksheet.Cells["D15"].Value = "Y";
                worksheet.Cells["D15"].Style.Font.Color = System.Drawing.Color.Green;
            }
            else
            {
                worksheet.Cells["D15"].Value = "N";
                worksheet.Cells["D15"].Style.Font.Color = System.Drawing.Color.Red;
            }
            if (SSBPOsysAuditResults.CPUaResult == "Pass")
            {
                worksheet.Cells["D16"].Value = "Y";
                worksheet.Cells["D16"].Style.Font.Color = System.Drawing.Color.Green;
            }
            else
            {
                worksheet.Cells["D16"].Value = "N";
                worksheet.Cells["D16"].Style.Font.Color = System.Drawing.Color.Red;
            }
            if (SSBPOsysAuditResults.InternetUpResult == "Pass")
            {
                worksheet.Cells["D18"].Value = "Y";
                worksheet.Cells["D18"].Style.Font.Color = System.Drawing.Color.Green;
            }
            else
            {
                worksheet.Cells["D18"].Value = "N";
                worksheet.Cells["D18"].Style.Font.Color = System.Drawing.Color.Red;
            }
            if (SSBPOsysAuditResults.InternetDownResult == "Pass")
            {
                worksheet.Cells["D19"].Value = "Y";
                worksheet.Cells["D19"].Style.Font.Color = System.Drawing.Color.Green;
            }
            else
            {
                worksheet.Cells["D19"].Value = "N";
                worksheet.Cells["D19"].Style.Font.Color = System.Drawing.Color.Red;
            }

            if (isTESTING)
            {
                ef.Save(@"\\Filesvr4\it\WinAudit\Test - Results_Archive\TESTING - " + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf");
                SSBPOsysAuditResults.attachmentFilename = @"\\Filesvr4\it\WinAudit\Test - Results_Archive\TESTING - " + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf";
            }
            else
            {
                ef.Save(@"\\Filesvr4\it\WinAudit\Results_Archive\" + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf");
                SSBPOsysAuditResults.attachmentFilename = @"\\Filesvr4\it\WinAudit\Results_Archive\" + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf";
            }

                //Save results as pdf
            
            ef = null;
        }
        private static SysAuditResults getAuditStatus(SysAuditResults SSBPOsysAuditResults)
        {

            if (SSBPOsysAuditResults.OSaResult == "Pass" & SSBPOsysAuditResults.CPUaResult == "Pass" & SSBPOsysAuditResults.RAMaResult == "Pass" & SSBPOsysAuditResults.InternetUpResult == "Pass" & SSBPOsysAuditResults.InternetDownResult == "Pass" & SSBPOsysAuditResults.HDDaResult == "Pass")
            {
                SSBPOsysAuditResults.aResult = "Pass";
            }
            else if (SSBPOsysAuditResults.aResult == "Pending")
            {
                SSBPOsysAuditResults.aResult = "Pending";
            }
            else
            {
                SSBPOsysAuditResults.aResult = "Fail";
            }

            if (SSBPOsysAuditResults.cCPU.Contains("given"))
            {
                SSBPOsysAuditResults.aFailedReason = "**Needs manual processing**";
            }

            if (!string.IsNullOrEmpty(SSBPOsysAuditResults.aFailedReason))
            {
                SSBPOsysAuditResults.aFailedReason = SSBPOsysAuditResults.aFailedReason.Substring(2);
            }

            return SSBPOsysAuditResults;
        }
        // Function create workbook to import results into Bit-Lever
        private static void createBitLeverImport(List<SysAuditResults> CandidatesList)
        {
            string file2Import = "";

            if (isTESTING)
            {
                file2Import = string.Format(@"\\filesvr4\IT\WinAudit\4BitLeverImport\TEST\BitLeverImport{0:yyyy-MM-dd_hh-mm-ss-tt}" + " Results.xls", DateTime.Now);
            }
            else
            {
                file2Import = string.Format(@"\\filesvr4\IT\WinAudit\4BitLeverImport\BitLeverImport{0:yyyy-MM-dd_hh-mm-ss-tt}" + " Results.xls", DateTime.Now);
            }

                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Bit-Lever Import");
                int index = 1;

                //Write the headers
                worksheet.Cells["A1"].Value = "Audit Run Date";
                worksheet.Cells["B1"].Value = "Candidate Name";
                worksheet.Cells["C1"].Value = "Candidate Email";
                worksheet.Cells["D1"].Value = "Notes";
                worksheet.Cells["E1"].Value = "SysAudit Status";
                worksheet.Cells["F1"].Value = "Processed By";
                worksheet.Cells["G1"].Value = "Results Sent?";
                worksheet.Cells["H1"].Value = "Fail Reason";
                worksheet.Cells["I1"].Value = "Date Processed";

            //write a row for each systemaudit result in the list
            foreach (SysAuditResults r in CandidatesList)
                {
                    worksheet.Cells[index, 0].Value = r.auditDate;
                    worksheet.Cells[index, 1].Value = r.cName;
                    worksheet.Cells[index, 2].Value = r.cEmail;
                    worksheet.Cells[index, 3].Value = r.aResultSummary;
                    worksheet.Cells[index, 4].Value = r.aResult;
                    worksheet.Cells[index, 5].Value = Environment.UserName;

                    if (r.aResult == "Pending")
                        worksheet.Cells[index, 6].Value = "No";
                    else
                    {
                        worksheet.Cells[index, 6].Value = "Yes";
                    }

                    worksheet.Cells[index, 7].Value = r.aFailedReason; 
                    worksheet.Cells[index, 8].Value = DateTime.Now;
                    index = index + 1;
                }

                workbook.Save(file2Import);
           

        }
        // Function to send results to candidates or helpdesk if the process fails
        //private static void sendResults(List<SysAuditResults> CandidatesList)
        //{
        //    // For each candidate results
        //    foreach (SysAuditResults r in CandidatesList)
        //    {
        //        //Check if there is a pdf file with the candidate's name
        //        string attachmentFilename = @"\\filesvr4\IT\WinAudit\Results_Archive\" + r.cName.Trim() + " SystemAudit Results.pdf";
        //        if (File.Exists(@"\\filesvr4\IT\WinAudit\Results_Archive\" + r.cName.Trim() + " SystemAudit Results.pdf"))
        //        {
        //            string s5 = r.cName + " - " + r.aResult + ". Sending results...";
        //            Console.SetCursorPosition((Console.WindowWidth - s5.Length) / 2, Console.CursorTop);
        //            Console.WriteLine(s5);

        //            sendMail(r.cEmail, attachmentFilename, r.cName, isTESTING);

        //            string s6 = r.cName + " results were sent.";
        //            Console.SetCursorPosition((Console.WindowWidth - s6.Length) / 2, Console.CursorTop);
        //            Console.WriteLine(s6);
        //        }
        //        else
        //        {
        //            //If not then send email to helpdesk
        //            sendErrorMail("brodriguez@statesidebpo.com", r.cName);
        //        }
        //    }

        //}
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
                    sysAuditResults.HDDaResult = "Fail";
                    sysAuditResults.aFailedReason = ", Hard drive insufficient";
                }
                else
                {
                    sysAuditResults.HDDaResult = "Pass";
                }
            }

            if (l.Contains("Operating"))
            {
                sysAuditResults.cOS = l.ToString().Substring(18);
                string cOSs = sysAuditResults.cOS.Substring(0, 10);
                string[] cOS1s = cOSs.Split(' ');
                string cOS2 = cOS1s[1];

                if (cOS2.Trim().Length > 1)
                {
                    cOS1s = cOS2.Split('.');
                    cOS2 = cOS1s[0];
                }


                if (Convert.ToInt64(cOS2.Trim()) < 7)
                {
                    sysAuditResults.OSaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", OS insufficient";
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
               // sysAuditResults.cRAM = l.ToString().Substring(54);
                string[] cRAMs = sysAuditResults.cRAM.Split('-');
                string cRAM1 = cRAMs[1].Replace("GB]", "").Replace("[Total RAM = ", "");

                if (Convert.ToInt32(cRAM1) < 2)
                {
                    sysAuditResults.RAMaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", RAM insufficient ";
                }
                else
                {
                    sysAuditResults.RAMaResult = "Pass";

                }
            }

            if (l.Contains("Network"))
            {

                sysAuditResults.cInternetUp = l.ToString().Substring(62).Replace("]", "").Trim();

                if (l.ToString().Substring(62).Replace("]", "").Trim().Contains("Kbps"))
                {
                    if (Convert.ToUInt32(sysAuditResults.cInternetUp) < 1000)
                    {
                        sysAuditResults.InternetUpResult = "Fail";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Upload speed insufficient";
                    }
                    else
                    {

                        sysAuditResults.InternetUpResult = "Pass";
                    }
                }
                else
                {

                    if (Convert.ToUInt32(sysAuditResults.cInternetUp) < 1)//could be problems
                    {
                        sysAuditResults.InternetUpResult = "Fail";

                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Upload speed insufficient";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Pass";

                    }

                 }

                sysAuditResults.cInternetDown = l.ToString().Substring(34, 10).Replace("]", "");


                if (Convert.ToUInt32(sysAuditResults.cInternetDown) < 3)
                {
                    sysAuditResults.InternetDownResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Download speed insufficient";
                }
                else
                {
                    sysAuditResults.InternetDownResult = "Pass";
                }



            }

            if (l.Contains("CPU Score "))
            {
                sysAuditResults.cCPU = l.ToString().Substring(26);

                string[] cCPUScores = sysAuditResults.cCPU.Split('=');
                string cCPUScore1 = cCPUScores[1].Replace(" ", "");
                cCPUScore1 = cCPUScores[1].Replace("] - [Processor", "").Trim();
                if (!cCPUScore1.Contains("given"))
                {
                    if (Convert.ToDouble(cCPUScore1) >= 4.8)
                    {
                        sysAuditResults.CPUaResult = "Pass";
                    }
                    else
                    {
                        sysAuditResults.CPUaResult = "Fail";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", CPU insufficient";

                    }
                }
                else
                {
                    sysAuditResults.CPUaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", **Needs manual processing**";
                    sysAuditResults.needsManualProcessing = true;
                    sysAuditResults.aResult = "Pending";
                }
            }

           
            sysAuditResults.aResultSummary = "[OS = " + sysAuditResults.cOS + "] - [C" + sysAuditResults.cCPU + " - " + sysAuditResults.cRAM + " - " + sysAuditResults.cHDD + " - [Download Speed = " + sysAuditResults.cInternetDown + " Mbps] - [Upload Speed = " + sysAuditResults.cInternetUp + " Mbps]";
            return sysAuditResults;
        }
        private static SysAuditResults getmManualResultsObject(SysAuditResults sysAuditResults, string l)
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

            if (l.Contains("Total Space"))
            {
                sysAuditResults.cHDD = l.ToString().Substring(27);

                string[] cHDDs = sysAuditResults.cHDD.Split('=');
                string cHDD1 = cHDDs[3].Replace("GB]", "");

                if (Convert.ToInt32(cHDD1) < 25)
                {
                    sysAuditResults.HDDaResult = "Fail";
                    sysAuditResults.aFailedReason = ", Hard drive insufficient";
                }
                else
                {
                    sysAuditResults.HDDaResult = "Pass";
                }
            }

            if (l.Contains("Operating"))
            {
                sysAuditResults.cOS = l.ToString().Substring(18);
                string cOSs = sysAuditResults.cOS.Substring(14, 2);
                //string[] cOS1s = cOSs.Split(' ');
                //string cOS2 = cOS1s[1];

                //if (cOS2.Trim().Length > 1)
                //{
                //    cOS1s = cOS2.Split('.');
                //    cOS2 = cOS1s[0];
                //}


                if (Convert.ToInt64(cOSs.Trim()) < 7)
                {
                    sysAuditResults.OSaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", OS insufficient";
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
        
            if (l.Contains("[Total RAM = "))
            {
                sysAuditResults.cRAM = l.ToString().Replace("RAM (Memory) Results: [Total RAM = ", "");

                // sysAuditResults.cRAM = l.ToString().Substring(54);
                string[] cRAMs = sysAuditResults.cRAM.Split(',');
                string cRAM1 = cRAMs[0].Replace("GB]", "");

                if (Convert.ToInt32(cRAM1) < 2)
                {
                    sysAuditResults.RAMaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", RAM insufficient ";
                }
                else
                {
                    sysAuditResults.RAMaResult = "Pass";

                }
            }

            if (l.Contains("Network"))
            {

                sysAuditResults.cInternetUp = l.ToString().Substring(62).Replace("]", "").Trim();

                if (l.ToString().Substring(62).Replace("]", "").Trim().Contains("Kbps"))
                {
                    if (Convert.ToUInt32(sysAuditResults.cInternetUp) < 1000)
                    {
                        sysAuditResults.InternetUpResult = "Fail";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Upload speed insufficient";
                    }
                    else
                    {

                        sysAuditResults.InternetUpResult = "Pass";
                    }
                }
                else
                {

                    if (Convert.ToUInt32(sysAuditResults.cInternetUp) < 1)//could be problems
                    {
                        sysAuditResults.InternetUpResult = "Fail";

                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Upload speed insufficient";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Pass";

                    }

                }

                sysAuditResults.cInternetDown = l.ToString().Substring(34, 10).Replace("]", "");


                if (Convert.ToUInt32(sysAuditResults.cInternetDown) < 3)
                {
                    sysAuditResults.InternetDownResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Download speed insufficient";
                }
                else
                {
                    sysAuditResults.InternetDownResult = "Pass";
                }



            }

            if (l.Contains("CPU Score "))
            {
                sysAuditResults.cCPU = l.ToString().Substring(26);

                string[] cCPUScores = sysAuditResults.cCPU.Split('=');
                string cCPUScore1 = cCPUScores[1].Replace(" ", "");
                cCPUScore1 = cCPUScores[1].Replace("] - [Processor", "").Trim().Replace("]","");
                if (!cCPUScore1.Contains("given"))
                {
                    if (Convert.ToDouble(cCPUScore1) >= 4.8)
                    {
                        sysAuditResults.CPUaResult = "Pass";
                    }
                    else
                    {
                        sysAuditResults.CPUaResult = "Fail";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", CPU insufficient";

                    }
                }
                else
                {
                    sysAuditResults.CPUaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", **Needs manual processing**";
                    sysAuditResults.needsManualProcessing = true;
                    sysAuditResults.aResult = "Pending";
                }
            }


            sysAuditResults.aResultSummary = "[OS = " + sysAuditResults.cOS + "] - [C" + sysAuditResults.cCPU + " - " + sysAuditResults.cRAM + " - " + sysAuditResults.cHDD + " - [Download Speed = " + sysAuditResults.cInternetDown + " Mbps] - [Upload Speed = " + sysAuditResults.cInternetUp + " Mbps]";
            return sysAuditResults;
        }
        //Get unseen emails from Sysaudit mailbox


        //Declare struct to hold the SysAudit 
        struct SysAuditResults
        {

            private bool _needsmanualprocessing;

            private string _afailedreason;
            private string _date;
            private string _name;
            private string _email;
            private string _host;
            private string _hdd;
            private string _cpu;
            private string _os;
            private string _ram;
            private string _internetUp;
            private string _internetDown;
            private string _results;
            private string _resultssummary;
            private string _attachmentfilename;

            private string _cpuresult;
            private string _osresults;
            private string _ramresults;
            private string _hddresults;
            private string _internetupresults;
            private string _internetdownresults;



            public bool needsManualProcessing
            {
                get
                {
                    return _needsmanualprocessing;
                }
                set
                {

                    _needsmanualprocessing = value;
                }
            }

            public string attachmentFilename
            {
                get
                {
                    return _attachmentfilename;
                }
                set
                {

                    _attachmentfilename = value;
                }
            }
            public string aFailedReason
            {
                get
                {
                    return _afailedreason;
                }
                set
                {

                    _afailedreason = value;
                }
            }
            public string auditDate
            {
                get
                {
                    return _date;
                }
                set
                {

                    _date = value;
                }
            }
            public string cName
            {
                get
                {
                    return _name;
                }
                set
                {

                    _name = value;
                }
            }
            public string cEmail
            {
                get
                {
                    return _email;
                }
                set
                {

                    _email = value;
                }
            }
            public string cHost
            {
                get
                {
                    return _host;
                }
                set
                {

                    _host = value;
                }
            }
            public string cHDD
            {
                get
                {
                    return _hdd;
                }
                set
                {

                    _hdd = value;
                }
            }
            public string cCPU
            {
                get
                {
                    return _cpu;
                }
                set
                {

                    _cpu = value;
                }
            }
            public string cOS
            {
                get
                {
                    return _os;
                }
                set
                {

                    _os = value;
                }
            }
            public string cRAM
            {
                get
                {
                    return _ram;
                }
                set
                {

                    _ram = value;
                }
            }
            public string cInternetUp
            {
                get
                {
                    return _internetUp;
                }
                set
                {

                    string[] tf = value.Split(' ');
                    int idx = 0;

                    if (tf[0] == "" | tf[0] == ":" | tf[0] == "=")
                        idx = 1;
                    if(tf[idx].ToString().Length > 2) { 
                        _internetUp = tf[idx].ToString().Substring(0, 3).Replace(".", "").Replace(";", "");
                    }
                    if(tf[idx].ToString().Length == 1) {
                        _internetUp = tf[idx].ToString();
                    }
                    else
                    {
                        _internetUp = tf[idx].ToString().Substring(0, 2).Replace(".", "").Replace(";", "");

                    }
                }
            }
            public string cInternetDown
            {
                get
                {
                    return _internetDown;
                }
                set
                {
                    string[] tf = value.Split(' ');
                    int idx = 0;

                    if (tf[0] == "" | tf[0] == ":")
                        idx = 1;

                    _internetDown = tf[idx].ToString().Substring(0, 2).Replace(".", "");

                }
            }
            public string aResult
            {
                get
                {
                    return _results;
                }
                set
                {

                    _results = value;
                }
            }
            public string aResultSummary
            {
                get
                {
                    return _resultssummary;
                }
                set
                {

                    _resultssummary = value;
                }
            }


            public string CPUaResult
            {
                get
                {
                    return _cpuresult;
                }
                set
                {

                    _cpuresult = value;
                }
            }
            public string OSaResult
            {
                get
                {
                    return _osresults;
                }
                set
                {

                    _osresults = value;
                }
            }
            public string RAMaResult
            {
                get
                {
                    return _ramresults;
                }
                set
                {

                    _ramresults = value;
                }
            }
            public string HDDaResult
            {
                get
                {
                    return _hddresults;
                }
                set
                {

                    _hddresults = value;
                }
            }
            public string InternetUpResult
            {
                get
                {
                    return _internetupresults;
                }
                set
                {

                    _internetupresults = value;
                }
            }
            public string InternetDownResult
            {
                get
                {
                    return _internetdownresults;
                }
                set
                {

                    _internetdownresults = value;
                }
            }

        }
        public static void sendMail(string recipient, string attachmentFilename, string cadidateName, bool isTESTING)
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

                    if (isTESTING)
                    {
                        to = new MailAddress("brodriguez@statesidebpo.com");
                    }

                    MailMessage myMail = new MailMessage(from, to);
                    myMail.IsBodyHtml = true;                 
                    myMail.Subject = "System audit results for " + cadidateName;

                    if (isTESTING)
                    {
                        myMail.Subject = "TESTING - System audit results for " + cadidateName;
                    }
                   
                   string body = @"<p style =""font-size:21px"">Dear " + cadidateName + ",<br><br>" + "This email is to inform you of your system audit results. Please see the attachment. If you have any technical questions regarding your results, please reach out to us via email at <a mailto:winaudit@statesidebpo.com>winaudit@statesidebpo.com</a>.</p>";
                   body = body + @"<p style =""font-size:18px""><i>(Please note: If you are unable to view the attachment, you may need to download and install Adobe Acrobat Reader DC or a similar program that allows the viewing of PDF documents)</i>";

                    Attachment inlineStatetSideLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\StatesideLogo.png");
                    Attachment inlineBitLeverLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\Bit-LeverLogo.png");

                    Attachment SysAuditResults = new Attachment(@"\\Filesvr4\it\WinAudit\Results_Archive\" + cadidateName + " SystemAudit Results.pdf");

                    if (isTESTING)
                    {
                        SysAuditResults = new Attachment(@"\\Filesvr4\it\WinAudit\\Test - Results_Archive\TESTING - " + cadidateName + " SystemAudit Results.pdf");
                    }

                    myMail.Attachments.Add(SysAuditResults);
                    myMail.Attachments.Add(inlineStatetSideLogo);
                    myMail.Attachments.Add(inlineBitLeverLogo);

                    inlineBitLeverLogo.ContentDisposition.Inline = true;
                    inlineStatetSideLogo.ContentDisposition.Inline = true;

                    inlineBitLeverLogo.ContentId = "BitLeverLogo";
                    inlineStatetSideLogo.ContentId =" StatetSideLogo";
                    SysAuditResults.ContentId = "Pdf";
                                      

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

                if (isTESTING) { 

                   to = new MailAddress("brodriguez@statesidebpo.com");
                }

                MailMessage myMail = new MailMessage(from, to);
                myMail.IsBodyHtml = true;
                myMail.Subject = DateTime.Now + " SystemAudit Processing run completed successfully";
               

                string bd = "";
                if (CandidatesList.Count() == 1)
                {
                    bd = "<h2> " + CandidatesList.Count() + " Audit was processed</h2>";
                }
                else
                {
                    bd = "<h2> " + CandidatesList.Count() + " Audits were processed</h2>";
                }

               
                bd = bd + "<table style = " + "' tr:nth-child(even) {background-color: #f2f2f2} '" +  " border = " + "1" + " border-radius= " + "10px" + " cellpadding = " + "6" + " cellspacing = " + "5" + "><tbody>";


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
                        sendMail(c.cEmail, c.attachmentFilename, c.cName, isTESTING);
                    }
                    else
                    {
                        if(c.aResult == "Fail") {
                            bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'43%'" + ">" + c.cEmail + "</td><td  width=" + "'43%'" + "><font color='red'>" + c.aResult + reason + "</font></td></tr>";
                        }
                        if(c.aResult == "Pass")
                        {                        
                            bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'28%'" + ">" + c.cEmail + "</td><td ' width=" + "'43%'" + "><font color='green'>" + c.aResult + reason + "</font></td></tr>";
                        }

                        if (c.aResult == "Pending")
                        {
                            bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'28%'" + ">" + c.cEmail + "</td><td  width=" + "'43%'" + "><font color='blue'>" + c.aResult + reason + "</font></td></tr>";
                        }

                        sendMail(c.cEmail, c.attachmentFilename, c.cName, isTESTING);
                    }

                }

                bd = bd + "</tbody></table>";

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
