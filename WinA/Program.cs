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

        public static bool isTESTING { get; private set; }

        static void Main(string[] args)
        {
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("################################################################################");
            Console.WriteLine("#######################      SysAuditPro  ver.2    #############################");
            Console.WriteLine("");
            // Declare app instance / list to hold all sysaudits 

            SystemAudit xxx_SysAudit = new SystemAudit();
            isTESTING = false;
            //isTESTING = true;
            int ProcessedEmails;
            int ProcessedManualEmails;

            List<SysAuditResults> CandidatesList = new List<SysAuditResults>();
            List<SysAuditResults> ManualCandidatesList = new List<SysAuditResults>();

            using (ImapClient imapClient = new ImapClient("secure.emailsrvr.com", "systemaudit@statesidebpo.com", "G4d5fG34!df!!2ew", AuthMethods.Login, 993, true))
            {

                Lazy<AE.Net.Mail.MailMessage>[] Manualmsgs = getManualMailMessages(imapClient);   //Need to improve this
                ProcessedManualEmails = proccessManualAudits(CandidatesList, imapClient, Manualmsgs);
               

                Lazy<AE.Net.Mail.MailMessage>[] msgs = getMailMessages(imapClient);
                ProcessedEmails = processNormalAudits(CandidatesList, imapClient, msgs);


                if (CandidatesList.Count() > 0)
                {
                    createBitLeverImport(CandidatesList);
                    sendCompletionNotification(CandidatesList);
                }


            }
            Console.WriteLine();
            Console.WriteLine("");
            Console.WriteLine("################################################################################");
            string s = ProcessedEmails + " Audits and " + ProcessedManualEmails + " manual Audits were processed. Good Bye!";
            Console.SetCursorPosition((Console.WindowWidth - s.Length) / 2, Console.CursorTop);
            Console.WriteLine(s);
            Console.WriteLine();
            Console.WriteLine("################################################################################");       
            Console.WriteLine();
            Thread.Sleep(5000);
        }

        private static int proccessManualAudits(List<SysAuditResults> CandidatesList, ImapClient imapClient, Lazy<AE.Net.Mail.MailMessage>[] Manualmsgs)
        {

            int ProcesseManualdEmails = 0;

            if (Manualmsgs.Count() != 0)
            {
                if (Manualmsgs.Count() == 1)
                {

                    string s9 = "There is " + Manualmsgs.Count() + " Manual SystemAudit to process.";
                    Console.SetCursorPosition((Console.WindowWidth - s9.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s9);
                 
                }
                else
                {
                    Console.WriteLine();
                    string s8 = "There are " + Manualmsgs.Count() + " Manual SystemAudits to process.";
                    Console.SetCursorPosition((Console.WindowWidth - s8.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s8);
                   
                }

                string sx = "Starting...";
                Console.SetCursorPosition((Console.WindowWidth - sx.Length) / 2, Console.CursorTop);
                Console.WriteLine(sx);
            

                HtmlToText stripHtml = new HtmlToText();

                //Foreach unseen email found in the mailbox
                foreach (Lazy<AE.Net.Mail.MailMessage> msg in Manualmsgs)
                {
                    //Declare sysaudit results object
                    SysAuditResults SSBPOsysAuditResults = new SysAuditResults();

                    // Flag each email as seen
                    imapClient.AddFlags(Flags.Seen, msg.Value);

                    ProcesseManualdEmails = ProcesseManualdEmails + 1;

                    string s1 = "Processing {0} of {1}.";
                    Console.SetCursorPosition((Console.WindowWidth - s1.Length) / 2, Console.CursorTop);
                    Console.Write(s1, ProcesseManualdEmails, Manualmsgs.Count());


                    if ((msg.Value.Body != "") && msg.Value != null)
                    {                        

                        string[] lines = stripHtml.Convert(msg.Value.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                        List<String> line = lines.ToList();

                        SysAuditXLWApp = new Excel.Application();
                        SysAuditWWorkBook = SysAuditXLWApp.Workbooks.Open(@"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xltm"); // Open the SysAudit Excel template
                        SysAuditWWorkSheet = SysAuditWWorkBook.Worksheets[1] as Excel.Worksheet; // Set sheet 1 as the active sheet in Excel template
                      //  SysAuditXLWApp.Visible = true;

                       
                        SysAuditWWorkSheet.Cells[5, 2] = "Computer Name: ";
                        SysAuditWWorkSheet.Cells[6, 2] = "Current User: ";                      

                        foreach (string l in line.ToList()) //Foreach of the lines in the email
                        {
                            if (l.ToString().Contains("Audit Run Date"))
                            {
                                string ln = "Audit Date Time: " + l.ToString().TrimStart().Replace("Audit Run Date", "");
                                SysAuditWWorkSheet.Cells[2, 2] = ln;
                                SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, ln);
                            }

                            if (l.ToString().Contains("Candidate Name"))
                            {
                                string ln = "Full Name: " + l.ToString().TrimStart().Replace("Candidate Name", "");
                                SysAuditWWorkSheet.Cells[3, 2] = ln;                                
                                SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, ln);
                            }

                            if (l.ToString().Contains("Candidate Email"))
                            {
                                string ln = "Email Address: " + l.ToString().TrimStart().Replace("Candidate Email", "");
                                SysAuditWWorkSheet.Cells[4, 2] = ln;                                
                                SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, ln);
                            }

                            if (l.ToString().Contains("Notes"))
                            {
                                string[] values = l.Split('-');
                                //Write the line value to the active sheet in Excel template
                                string trn = "";
                                string rsn = "";
                                string sRamc = "";

                                string hdds = "";
                                string hddsp = "";
                                string hddas = "";
                                string shdd = "";

                                string cpus = "";
                                string cpun = "";
                                string sCpuc = "";

                                string manualUpSpeed = "";
                                string manualDownSpeed = "";
                                string internetSpeed = "";                               

                                foreach (string t in values)
                                {
                                    string v = t.Replace("Notes", "");

                                    if (v.Contains("Windows"))
                                    {
                                        string ln = "Operating System: " + v.ToString().TrimStart().Replace("[OS = ", "").Replace("]", "");
                                        SysAuditWWorkSheet.Cells[7, 2] = ln;
                                        SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, ln);
                                    }

                                    if (v.Contains("CPU Score"))
                                    {
                                        cpus = v.ToString().TrimStart();
                                    }

                                    if (v.Contains("Processor"))
                                    {
                                        cpun = v.ToString().TrimStart();
                                        sCpuc = "CPU (Processor) Results: " + cpus + " - " + cpun + "]".Replace("CCPU", "CPU");
                                        SysAuditWWorkSheet.Cells[8, 2] = sCpuc;
                                        SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, sCpuc);
                                    }
                                    if (v.Contains("RAM Score"))
                                    {
                                        rsn = v.ToString().TrimStart();
                                    }

                                    if (v.Contains("Total RAM"))
                                    {
                                        trn =  v.ToString().TrimStart();
                                        sRamc = "RAM (Memory) Results: " + rsn + " - " + trn;
                                        SysAuditWWorkSheet.Cells[9, 2] = sRamc;
                                        SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, sRamc);
                                    }

                                    if (v.Contains("Disk Score"))
                                    {
                                        hdds = v.ToString().TrimStart();
                                    }
                                    if (v.Contains("Total Space"))
                                    {
                                        hddsp = v.ToString().TrimStart();
                                    }
                                    if (v.Contains("Available Space"))
                                    {
                                        hddas = v.ToString().TrimStart();
                                        shdd = "Disk (Hard Drive) Results:  " + hdds + " - " + hddsp + " - " + hddas;
                                        SysAuditWWorkSheet.Cells[10, 2] = shdd;
                                        SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, shdd);
                                    }
                                    if (v.Contains("Download"))
                                    {
                                        manualDownSpeed = v.ToString().TrimStart();
                                    }

                                    if (v.Contains("Upload"))
                                    {
                                        manualUpSpeed = v.ToString().TrimStart();
                                        internetSpeed = "Network Results: " + manualDownSpeed.Trim().Replace(" = ",": ") + " - " + manualUpSpeed.Trim().Replace(" = ", ": ");
                                        SysAuditWWorkSheet.Cells[11, 2] = internetSpeed;
                                        SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, internetSpeed);
                                    }
                                }
                            }
                        }


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
                        
                        if (!SSBPOsysAuditResults.cCPU.Contains("given"))
                        {
                            SysAuditXLWApp.Run("Sheet2.SaveAsC");  //Run SaveAsC macro on the Excel template to export results to pdf                                 // 
                        }

                        if (SSBPOsysAuditResults.cCPU.Contains("given"))
                        {
                            SSBPOsysAuditResults.aFailedReason = "**Needs manual processing**";
                        }

                        if (!string.IsNullOrEmpty(SSBPOsysAuditResults.aFailedReason))
                        {
                            SSBPOsysAuditResults.aFailedReason = SSBPOsysAuditResults.aFailedReason.Substring(2);
                        }
                       
                        object misValue = System.Reflection.Missing.Value; //Get misssing.vlaue variable
                        SysAuditWWorkBook.Close(false, misValue, misValue);
                        SysAuditWWorkSheet = null;
                        SysAuditWWorkBook = null;
                        SysAuditXLWApp.Quit();
                    }

                    CandidatesList.Add(SSBPOsysAuditResults);

                }

              //  sendCompletionNotification(CandidatesList);

            }

            return ProcesseManualdEmails;

        }

        private static int processNormalAudits(List<SysAuditResults> CandidatesList, ImapClient imapClient, Lazy<AE.Net.Mail.MailMessage>[] msgs)
        {

            killExcel();
            
            int ProcessedEmails = 0;

            if (msgs.Count() != 0)
            {
                if (msgs.Count() == 1)
                {
                    string s1 = "There is " + msgs.Count() + " SystemAudit to process.";
                    Console.SetCursorPosition((Console.WindowWidth - s1.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s1);
                    Console.WriteLine();
                }
                else
                {

                    string s2 = "There are " + msgs.Count() + " SystemAudits to process.";
                    Console.SetCursorPosition((Console.WindowWidth - s2.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s2);
                    Console.WriteLine();

                }

                string s3 = "Starting...";
                Console.SetCursorPosition((Console.WindowWidth - s3.Length) / 2, Console.CursorTop);
                Console.WriteLine(s3);

                HtmlToText stripHtml = new HtmlToText();
                //Foreach unseen email found in the mailbox
                foreach (Lazy<AE.Net.Mail.MailMessage> msg in msgs)
                {
                    ProcessedEmails = ProcessedEmails + 1;
                    //Declare sysaudit results object
                    SysAuditResults SSBPOsysAuditResults = new SysAuditResults();

                    // Flag each email as seen
                    imapClient.AddFlags(Flags.Seen, msg.Value);

                    string s5 = "Processing {0} of {1}.";
                    Console.SetCursorPosition((Console.WindowWidth - s5.Length) / 2, Console.CursorTop);
                    Console.Write(s5, ProcessedEmails, msgs.Count());


                    if ((msg.Value.Body != "") && msg.Value != null)
                    {  
                        string[] lines = stripHtml.Convert(msg.Value.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                        int i = 2;
                        List<String> line = lines.ToList();

                        SysAuditXLWApp = new Excel.Application();
                        SysAuditWWorkBook = SysAuditXLWApp.Workbooks.Open(@"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xltm"); // Open the SysAudit Excel template
                        SysAuditWWorkSheet = SysAuditWWorkBook.Worksheets[1] as Excel.Worksheet; // Set sheet 1 as the active sheet in Excel template

                        foreach (string l in line.ToList()) //Foreach of the lines in the email
                        {
                            if (l.ToString() != "" & !l.ToString().Contains("Simplified Audit Results") & !l.ToString().Contains("www") & !l.Contains("Avast") & !l.Contains("antivirus")) //If the line is not empty or has unwanted text
                            {
                                SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, l); // Set the results object varialble to the value on the specified in the line 
                                SysAuditWWorkSheet.Cells[i, 2] = l.ToString().TrimStart(); //Write the line value to the active sheet in Excel template
                                i++;
                            }
                        }

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


                        if (!SSBPOsysAuditResults.cCPU.Contains("given"))
                        {
                            SysAuditXLWApp.Run("Sheet2.SaveAsC");  //Run SaveAsC macro on the Excel template to export results to pdf                                 // 
                        }

                        if (SSBPOsysAuditResults.cCPU.Contains("given"))
                        {
                            SSBPOsysAuditResults.aFailedReason = "**Needs manual processing**";
                        }

                        if (!string.IsNullOrEmpty(SSBPOsysAuditResults.aFailedReason))
                        {
                            SSBPOsysAuditResults.aFailedReason = SSBPOsysAuditResults.aFailedReason.Substring(2);
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

            return ProcessedEmails;
        }

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

                SysAuditWWorkSheet2.Cells[1, 1] = "Audit Run Date";
                SysAuditWWorkSheet2.Cells[1, 2] = "Candidate Name";
                SysAuditWWorkSheet2.Cells[1, 3] = "Candidate Email";
                SysAuditWWorkSheet2.Cells[1, 4] = "Notes";
                SysAuditWWorkSheet2.Cells[1, 5] = "SysAudit Status";
                SysAuditWWorkSheet2.Cells[1, 6] = "Processed By";
                SysAuditWWorkSheet2.Cells[1, 7] = "Results Sent?";
                SysAuditWWorkSheet2.Cells[1, 8] = "Fail Reason";
                SysAuditWWorkSheet2.Cells[1, 9] = "Date Processed";

                foreach (SysAuditResults r in CandidatesList)
                {
                    string day = r.auditDate;
                    day = day.Replace("p.m.", "PM");

                    SysAuditWWorkSheet2.Cells[index, 1] = day;
                    SysAuditWWorkSheet2.Cells[index, 2] = r.cName;
                    SysAuditWWorkSheet2.Cells[index, 3] = r.cEmail;
                    SysAuditWWorkSheet2.Cells[index, 4] = r.aResultSummary;
                    SysAuditWWorkSheet2.Cells[index, 5] = r.aResult;
                    SysAuditWWorkSheet2.Cells[index, 6] = Environment.UserName;
                    if (r.aResult == "Pending")
                        SysAuditWWorkSheet2.Cells[index, 7] = "No";
                    else
                    {
                        SysAuditWWorkSheet2.Cells[index, 7] = "Yes";
                    }
                    SysAuditWWorkSheet2.Cells[index, 8] = r.aFailedReason;
                    SysAuditWWorkSheet2.Cells[index, 9] = DateTime.Now;
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
            //try
            //{


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
                    string[] cRAMs = sysAuditResults.cRAM.Split('=');
                    string cRAM1 = cRAMs[2].Replace("GB]", "");

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
                    string[] uSpeed = l.Split('-');
                 

                    sysAuditResults.cInternetUp = uSpeed[1].ToString().Substring(16, 4).Replace(".","");

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
                    int Iups = Convert.ToInt32(sysAuditResults.cInternetUp.Substring(0,1));

                        if (Iups < 1)//could be problems
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

                if (l.Contains("CPU"))
                {
                    sysAuditResults.cCPU = "[" + l.ToString().Substring(26).Replace("CC", "C");

                    string[] cCPUScores = sysAuditResults.cCPU.Split('=');
                    string cCPUScore1 = cCPUScores[1].Replace("]  - [Processor", "").Trim();
                    cCPUScore1 = cCPUScore1.Replace("] - [Processor", "");


                    if (!cCPUScore1.Contains("given"))
                    {
                        decimal cCPU = Convert.ToDecimal(cCPUScore1);

                        if (cCPU >= Convert.ToDecimal(4.8))
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

                
            //}
            //catch (System.Exception ex)
            //{
            //    throw new ApplicationException
            //      ("Exception has occured: " + ex.Message);
            //}

            //if (!string.IsNullOrEmpty(sysAuditResults.aFailedReason))
            //{
            //    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason.Remove(0, 1);
            //}

            sysAuditResults.aResultSummary = "[OS = " + sysAuditResults.cOS + "] - " + sysAuditResults.cCPU + " - " + sysAuditResults.cRAM + " - " + sysAuditResults.cHDD + " - [Download Speed = " + sysAuditResults.cInternetDown + " Mbps] - [Upload Speed = " + sysAuditResults.cInternetUp + " Mbps]";
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

        private static Lazy<AE.Net.Mail.MailMessage>[] getManualMailMessages(ImapClient imapClient)
        {
            imapClient.SelectMailbox("INBOX");
            Regex regex = new Regex(@":");
            // Lazy<AE.Net.Mail.MailMessage>[] messages = imapClient.SearchMessages(SearchCondition.From("systemaudit@bit-lever.com"), false);
            Lazy<AE.Net.Mail.MailMessage>[] messages = imapClient.SearchMessages(SearchCondition.From("no-reply@bit-lever.com").And(SearchCondition.Unseen()));

            return messages;
        }

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

                   
                        _internetUp = value;
                  
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

                    if (tf[idx].Length > 1) { 
                    _internetDown = tf[idx].ToString().Substring(0, 2).Replace(".", "");
                    }
                    else
                    {
                        _internetDown = tf[idx].ToString();
                    }
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

        public static void sendMail(string recipient, string attachmentFilename, string cadidateName, bool isTESTING)
        {
            if (File.Exists(attachmentFilename))
            {

                try
                {
                    Outlook.Application otApp = new Outlook.Application();// create outlook object
                    Outlook.NameSpace ns = otApp.Session;

                    Outlook.Folder folder = otApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts) as Outlook.Folder;
                    Outlook.MailItem otMsg = otApp.CreateItemFromTemplate(@"\\filesvr4\IT\WinAudit\SysAudit App\System audit results.oft", folder) as Outlook.MailItem;///);Outlook.MailItem)otApp.CreateItem(Outlook.OlItemType.olMailItem); // Create mail object

                    otMsg.SendUsingAccount = getAccountForEmailAddress(otApp, "systemaudit@statesidebpo.com");
                    Outlook.Inspector oInspector = otMsg.GetInspector;

                    String sSource = "";
                    String sDisplayName = "";

                    if (isTESTING)
                    {
                        Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add("brodriguez@statesidebpo.com"); //jreiner@statesidebpo.com  brodriguez@statesidebpo.com
                        otMsg.BCC = "brodriguez@statesidebpo.com";
                        otMsg.Subject = "TESTING - SSBPO System audit results";
                        sDisplayName = "TESTING - " + cadidateName + " SystemAudit Results.pdf";
                        sSource = attachmentFilename;
                    }
                    else
                    {
                        Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add(recipient);                        
                        otMsg.BCC = "recruiters@statesidebpo.com";
                        otMsg.Subject = "SSBPO System audit results";
                        sDisplayName = cadidateName + " SystemAudit Results.pdf";
                        sSource = attachmentFilename;
                    }

                    otMsg.Recipients.ResolveAll();// validate recipient address                    

                    int iPos = (int)otMsg.Body.Length + 1;
                    int iAttType = (int)Outlook.OlAttachmentType.olByValue;
                    Outlook.Attachment oAttach = otMsg.Attachments.Add(sSource, iAttType, iPos, sDisplayName); // add attachment
                    otMsg.Save();
                    otMsg.Send();
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
                if (isTESTING)
                {
                    Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add("brodriguez@statesidebpo.com"); //brodriguez@statesidebpo.com jreiner  //helpdesk@statesidebpo.com
                    otRecip.Resolve();// validate recipient address
                    otMsg.Subject = "TESTING - " + DateTime.Now + " SystemAudit Processing run completed successfully";
                }
                else
                {
                    Outlook.Recipient otRecip = (Outlook.Recipient)otMsg.Recipients.Add("helpdesk@statesidebpo.com"); //brodriguez@statesidebpo.com jreiner  //helpdesk@statesidebpo.com
                    Outlook.Recipient recipBcc = otMsg.Recipients.Add("recruiters@statesidebpo.com");
                    recipBcc.Type = (int)Outlook.OlMailRecipientType.olBCC;
                    otRecip.Resolve();// validate recipient address
                    otMsg.Subject = DateTime.Now + " SystemAudit Processing run completed successfully";
                }

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

                    if (c.needsManualProcessing)
                    {

                        bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'33%'" + ">" + c.cEmail + "</td><td width=" + "'43%'" + ">" + c.aResult + reason + "</td></tr>";
                        sendMail(c.cEmail, attachmentFilename, c.cName, isTESTING);
                    }
                    else
                    {
                        bd = bd + "<tr><td  width=" + "'23%'" + ">" + c.cName + "</td><td width=" + "'33%'" + ">" + c.cEmail + "</td><td width=" + "'43%'" + ">" + c.aResult + reason + "</td></tr>";
                        sendMail(c.cEmail, attachmentFilename, c.cName, isTESTING);
                    }

                }

                bd = bd + "</tbody></table>";

                otMsg.HTMLBody = Regex.Replace(bd, @"[^\u0000-\u007F]", " ");


                otMsg.Send();

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

