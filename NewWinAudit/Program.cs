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
using System.Globalization;
using System.Data.Odbc;

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

            isTESTING = false;
            //isTESTING = true;

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("################################################################################");
            Console.WriteLine("#####################         WinAuditPro ver.3       ##########################");
            Console.WriteLine("################################################################################");
            Console.WriteLine();

            try
            {
                // Declare app instance / list to hold all sysaudits 
                SystemAudit xxx_SysAudit = new SystemAudit();
                List<SysAuditResults> CandidatesList = new List<SysAuditResults>();


                using (S22.Imap.ImapClient imapClient = new S22.Imap.ImapClient("secure.emailsrvr.com", 993, "systemaudit@statesidebpo.com", "G4xxsfsdf$$$-fsdfxG3ffsdafsdfa4!df!!2eX", AuthMethod.Login, true))
                {
                    string s = "Checking SystemAudits mailbox...";
                    Console.SetCursorPosition((Console.WindowWidth - s.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s);
                    Console.WriteLine("");

                    // Check mailbox and get any messages not seen and sent by systemaudit@bit-lever.com
                    IEnumerable<uint> uids = imapClient.Search(S22.Imap.SearchCondition.Unseen().And(S22.Imap.SearchCondition.From("systemaudit@bit-lever.com")));
                    IEnumerable<System.Net.Mail.MailMessage> messages = imapClient.GetMessages(uids);


                    IEnumerable<uint> manualuids = imapClient.Search(S22.Imap.SearchCondition.Unseen().And(S22.Imap.SearchCondition.From("no-reply@bit-lever.com")));
                    IEnumerable<System.Net.Mail.MailMessage> manualmessages = imapClient.GetMessages(manualuids);

                    Thread.Sleep(1000);

                    string s2 = "New Audits to process:  " + messages.Count();
                    Console.SetCursorPosition((Console.WindowWidth - s2.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s2);
                    xxx_SysAudit.ProcessedEmails = processNormalAudits(CandidatesList, messages);
                    Thread.Sleep(1000);

                    Console.WriteLine("");
                    Console.WriteLine("");

                    string sx = "Manual Audits to process:  " + manualmessages.Count();
                    Console.SetCursorPosition((Console.WindowWidth - sx.Length) / 2, Console.CursorTop);
                    Console.WriteLine(sx);
                    xxx_SysAudit.ProcessedManualEmails = proccessManualAudits(CandidatesList, manualmessages);
                    Thread.Sleep(1000);

                    sendCompletionNotification(CandidatesList);
                    sendToQuickBase(CandidatesList);

                    if (CandidatesList.Count() > 0)
                    {
                        createFileImport(CandidatesList);
                    }

                    Console.WriteLine("");
                    string s5 = "Processing completed. " + xxx_SysAudit.ProcessedEmails + " New Audits and " + xxx_SysAudit.ProcessedManualEmails + " Manual were processed. Good bye!";
                    Console.WriteLine("");
                    Console.SetCursorPosition((Console.WindowWidth - s5.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s5);
                    Console.WriteLine("");
                    Console.WriteLine("################################################################################");
                    Thread.Sleep(10000);
                }


            }
            catch (System.Exception err)
            {
                throw new CustomException("Processing failed: ", err);
            }

        }

        private static int processNormalAudits(List<SysAuditResults> CandidatesList, IEnumerable<MailMessage> messages)
        {
            try
            {

                int ProcessedEmails = 0;

                if (messages.Count() != 0)
                {
                    string s4 = "Processing new Audits...";
                    Console.SetCursorPosition((Console.WindowWidth - s4.Length) / 2, Console.CursorTop);
                    Console.WriteLine(s4);
                    Thread.Sleep(2000);

                    foreach (System.Net.Mail.MailMessage msg in messages)
                    {
                        ProcessedEmails = ProcessedEmails + 1;
                        string sz = "Processing {0} of {1}.";
                        Console.SetCursorPosition((Console.WindowWidth - sz.Length) / 2, Console.CursorTop);
                        Console.Write(sz, ProcessedEmails, messages.Count());

                        //Load the WinAudit workbook
                        string fileName = @"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xlsm";
                        var ef = ExcelFile.Load(fileName);

                        //Instanciate SystemAudit Results object
                        SysAuditResults SSBPOsysAuditResults = new SysAuditResults();
                        HtmlToText stripHtml = new HtmlToText();

                        // If message body is not empty 
                        if (msg.Body != "")
                        {
                            //Strip Html from emal and split into lines  
                            string[] line = stripHtml.Convert(msg.Body).Split(new string[] { "\r\n", "\n", "* " }, StringSplitOptions.None);

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
                }

                return ProcessedEmails;

            }
            catch (System.Exception err)
            {
                throw new CustomException("processNormalAudits Error: ", err);
            }



        }
        private static int proccessManualAudits(List<SysAuditResults> CandidatesList, IEnumerable<MailMessage> Manualmsgs)
        {

            try
            {

                int ProcesseManualdEmails = 0;

                if (Manualmsgs.Count() != 0)
                {
                    string sx = "Processing manual Audits...";
                    Console.SetCursorPosition((Console.WindowWidth - sx.Length) / 2, Console.CursorTop);
                    Console.WriteLine(sx);

                    HtmlToText stripHtml = new HtmlToText();

                    //Foreach unseen email found in the mailbox
                    foreach (System.Net.Mail.MailMessage msg in Manualmsgs)
                    {
                        //imapClient.RemoveMessageFlags(msg, null, MessageFlag.Seen);
                        ProcesseManualdEmails = ProcesseManualdEmails + 1;

                        string s1 = "Processing {0} of {1}.";
                        Console.SetCursorPosition((Console.WindowWidth - s1.Length) / 2, Console.CursorTop);
                        Console.Write(s1, ProcesseManualdEmails, Manualmsgs.Count());

                        //Instanciate SystemAudit Results object
                        SysAuditResults SSBPOsysAuditResults = new SysAuditResults();

                        //Load the WinAudit workbook
                        string fileName = @"\\filesvr4\IT\WinAudit\SysAudit App\WinAuditPro.xlsm";
                        var ef = ExcelFile.Load(fileName);


                        if (msg.Body != "")
                        {

                            string[] lines = stripHtml.Convert(msg.Body).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                            List<String> line = lines.ToList();

                            ExcelWorksheet worksheet = ef.Worksheets.ActiveWorksheet;

                            foreach (string l in line.ToList()) //Foreach of the lines in the email
                            {
                                if (l.ToString().Contains("Audit Run Date"))
                                {
                                    string ln = "Audit Date Time: " + l.ToString().TrimStart().Replace("Audit Run Date", "");
                                    SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, ln);
                                }

                                if (l.ToString().Contains("Candidate Name"))
                                {
                                    string ln = "Full Name: " + l.ToString().TrimStart().Replace("Candidate Name", "");
                                    SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, ln);
                                }

                                if (l.ToString().Contains("Candidate Email"))
                                {
                                    string ln = "Email Address: " + l.ToString().TrimStart().Replace("Candidate Email", "");
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
                                            SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, sCpuc);
                                        }
                                        if (v.Contains("RAM Score"))
                                        {
                                            rsn = v.ToString().TrimStart();
                                        }

                                        if (v.Contains("Total RAM"))
                                        {
                                            trn = v.ToString().TrimStart();
                                            sRamc = "RAM (Memory) Results: " + rsn + " - " + trn;
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
                                            SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, shdd);
                                        }
                                        if (v.Contains("Download"))
                                        {
                                            manualDownSpeed = v.ToString().TrimStart();
                                        }

                                        if (v.Contains("Upload"))
                                        {
                                            manualUpSpeed = v.ToString().TrimStart();
                                            internetSpeed = "Network Results: " + manualDownSpeed.Trim().Replace(" = ", ": ") + " - " + manualUpSpeed.Trim().Replace(" = ", ": ");
                                            SSBPOsysAuditResults = getResultsObject(SSBPOsysAuditResults, internetSpeed);
                                        }
                                    }//each
                                }
                            }
                        }

                        SSBPOsysAuditResults = getAuditStatus(SSBPOsysAuditResults);
                        generatePdf(ref ef, ref SSBPOsysAuditResults);
                        CandidatesList.Add(SSBPOsysAuditResults);
                    }
                }

                return ProcesseManualdEmails;

            }
            catch (System.Exception err)
            {
                throw new CustomException("processManualAudits Error: ", err);
            }
        }
        private static void generatePdf(ref ExcelFile ef, ref SysAuditResults SSBPOsysAuditResults)
        {
            //Fill out the candidates results worksheet and save as pdf to send out
            DateTime dt = Convert.ToDateTime(SSBPOsysAuditResults.auditDate);

            ExcelWorksheet worksheet = ef.Worksheets.ActiveWorksheet;
            worksheet.Cells["C5"].Value = SSBPOsysAuditResults.cName;


            worksheet.Cells["C6"].Value = dt;
            worksheet.Cells["C6"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
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

            string[] cHDD = SSBPOsysAuditResults.cHDD.Split('=');

            worksheet.Cells["C14"].Value = cHDD[3].Replace("]", "").Trim();

            //worksheet.Cells["C14"].Value = SSBPOsysAuditResults.cHDD.Substring(64).Replace("]", "").Replace("= ", "").Trim();


            string[] cRAM = SSBPOsysAuditResults.cRAM.Split('=');
            worksheet.Cells["C15"].Value = cRAM[2].Replace("]", "").Trim();

            if (SSBPOsysAuditResults.CPUaResult == "Pass" && SSBPOsysAuditResults.cCPU.Contains("given"))
            {
                worksheet.Cells["C16"].Value = "5.0";
            }
            else
            {
                worksheet.Cells["C16"].Value = SSBPOsysAuditResults.cCPU.Substring(12, 4).Replace("=", "").Trim().Replace("]", "");
            }



            
            worksheet.Cells["C18"].Value = SSBPOsysAuditResults.cInternetUp.Trim() + " Mbps";
            worksheet.Cells["C19"].Value = SSBPOsysAuditResults.cInternetDown.Trim() + " Mbps";

            if (Convert.ToDecimal(SSBPOsysAuditResults.cInternetUp) > 99)
            {
                worksheet.Cells["C18"].Value = SSBPOsysAuditResults.cInternetUp.Trim() + " Kbps";
            }
            if (Convert.ToDecimal(SSBPOsysAuditResults.cInternetDown) > 99)
            {
                worksheet.Cells["C19"].Value = SSBPOsysAuditResults.cInternetDown.Trim() + " Kbps";
            }

            worksheet.Cells["D13"].Value = SSBPOsysAuditResults.OSaResult;
            worksheet.Cells["D14"].Value = SSBPOsysAuditResults.HDDaResult.Replace("= ", "").Trim();
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

            if (SSBPOsysAuditResults.needsManualProcessing == false)
            {

                if (isTESTING == true)
                {
                    ef.Save(@"\\Filesvr4\it\WinAudit\Test - Results_Archive\TESTING - " + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf");
                    SSBPOsysAuditResults.attachmentFilename = @"\\Filesvr4\it\WinAudit\Test - Results_Archive\TESTING - " + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf";
                }
                else
                {
                    ef.Save(@"\\Filesvr4\it\WinAudit\Results_Archive\" + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf");
                    SSBPOsysAuditResults.attachmentFilename = @"\\Filesvr4\it\WinAudit\Results_Archive\" + SSBPOsysAuditResults.cName + " SystemAudit Results.pdf";
                }
            }
            //Save results as pdf

            ef = null;
        }
        private static SysAuditResults getAuditStatus(SysAuditResults SSBPOsysAuditResults)
        {

            if (SSBPOsysAuditResults.OSaResult == "Pass" & SSBPOsysAuditResults.CPUaResult == "Pass" & SSBPOsysAuditResults.RAMaResult == "Pass" & SSBPOsysAuditResults.InternetUpResult == "Pass" & SSBPOsysAuditResults.InternetDownResult == "Pass" & SSBPOsysAuditResults.HDDaResult == "Pass")
            {
                SSBPOsysAuditResults.aResult = "Pass";
                SSBPOsysAuditResults.resultSent = "Yes";
            }
            else if (SSBPOsysAuditResults.aResult == "Pending")
            {
                SSBPOsysAuditResults.aResult = "Pending";
                SSBPOsysAuditResults.resultSent = "No";
            }
            else
            {
                SSBPOsysAuditResults.aResult = "Fail";
                SSBPOsysAuditResults.resultSent = "Yes";
            }

            //if (SSBPOsysAuditResults.cCPU.Contains("given"))
            //{
            //    SSBPOsysAuditResults.aFailedReason = "**Needs manual processing**";
            //}

            if (!string.IsNullOrEmpty(SSBPOsysAuditResults.aFailedReason))
            {
                SSBPOsysAuditResults.aFailedReason = SSBPOsysAuditResults.aFailedReason.Substring(2);
            }

            return SSBPOsysAuditResults;
        }

        private static void createFileImport(List<SysAuditResults> CandidatesList)
        {
            string path = @"\\filesvr4\it\WinAudit\4BitLeverImport\AuditMasterLIVE.csv";

            if (!isTESTING) {
                if (File.Exists(path))
                {
                    foreach (SysAuditResults c in CandidatesList)
                    {
                        if (c.IsManual != true)
                        {
                            string line = Convert.ToDateTime(c.auditDate).ToString("g", System.Globalization.CultureInfo.CreateSpecificCulture("en-us")) + "," + c.cName.Replace('Â', ' ') + "," + c.cEmail + "," + c.aResultSummary + "," + c.aResult + "," + Environment.UserName + "," + c.resultSent + ", " + c.aFailedReason + ", " + DateTime.Now;
                            string createText = line;
                            File.AppendAllText(path, createText + Environment.NewLine, Encoding.UTF32);
                        }
                    }

                }
            }
        }


        private static void sendToQuickBase(List<SysAuditResults> CandidatesList)
        {
            try
            {
                foreach (SysAuditResults c in CandidatesList)
                {
                    OdbcConnection DbConnection = new OdbcConnection("DSN=QuickBase via QuNect user");
                    DbConnection.Open();

                    string insert = "insert into bmrksgqsn (Audit Run Date, Candidate Name, Candidate Email, SysAudit Status, Notes, Fail Reason) values(?,?,?,?,?,?)";
                    OdbcCommand commmand = new OdbcCommand(insert, DbConnection);
                    OdbcDataReader reader;

                    string fr = "N/A";

                


                    commmand.Parameters.AddWithValue("@Audit Run Date", OdbcType.DateTime).Value = Convert.ToDateTime(c.auditDate);

                    commmand.Parameters.AddWithValue("@Candidate Name", OdbcType.VarChar).Value = c.cName;

                    commmand.Parameters.AddWithValue("@Candidate Email", OdbcType.VarChar).Value = c.cEmail;

                    commmand.Parameters.AddWithValue("@SysAudit Status", OdbcType.VarChar).Value = c.aResult;
                
                    commmand.Parameters.AddWithValue("@Notes", OdbcType.VarChar).Value =  c.aResultSummary.Replace("(","[").Replace(")","]");

                if (c.aFailedReason == null)
                {
                        commmand.Parameters.AddWithValue("@Fail Reason", OdbcType.VarChar).Value = fr;

                    }
                    else
                    {
                       
                        commmand.Parameters.AddWithValue("@Fail Reason", OdbcType.VarChar).Value = c.aFailedReason;
                    }

                    reader = commmand.ExecuteReader();
                    DbConnection.Close();
                    Console.WriteLine("Completed import");

                }


            }

            catch (System.Exception ex)
            {
                throw new ApplicationException
                  ("Sending to DB Error: " + ex.Message);
            }
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
                    sysAuditResults.aFailedReason = "/ Hard drive insufficient";
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
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "/ OS insufficient";
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

            if (l.Contains(" RAM"))
            {
              
                sysAuditResults.cRAM = l.ToString().Substring(22);
                string[] cRAMs = sysAuditResults.cRAM.Split('=');
                string cRAM1 = cRAMs[2].Replace("GB]", "");

                if (Convert.ToInt32(cRAM1) < 2)
                {
                    sysAuditResults.RAMaResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + "/ RAM insufficient ";
                }
                else
                {
                    sysAuditResults.RAMaResult = "Pass";
                }
            }

            if (l.Contains("Network"))
            {

                string[] dp = null;
                string[] up = null;
                string[] uSpeed = l.Split('-');

                up = uSpeed[1].Split(':');
                sysAuditResults.cInternetUp = up[1].Trim().Replace("Mbps]", "").Replace("Kbps]", "");

                dp = uSpeed[0].Split(':');
                sysAuditResults.cInternetDown = dp[2].Trim().Replace("Mbps]", "").Replace("Kbps]", "");

                if (up[1].ToString().Trim().Contains("Kbps"))
                {
                    if (Convert.ToUInt32(sysAuditResults.cInternetUp.Substring(0, 1)) < 1000)
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
                    if (Convert.ToUInt32(sysAuditResults.cInternetUp.Substring(0, 1)) < 1)
                    {
                        sysAuditResults.InternetUpResult = "Fail";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Upload speed insufficient";
                    }
                    else
                    {
                        sysAuditResults.InternetUpResult = "Pass";
                    }
                }




                if (dp[2].ToString().Trim().Contains("Kbps"))
                {

                    sysAuditResults.InternetDownResult = "Fail";
                    sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", Download speed insufficient";

                }
                else
                {

                    if (sysAuditResults.cInternetDown.Contains("."))
                    {
                        string[] dIp = sysAuditResults.cInternetDown.Split('.');
                        sysAuditResults.cInternetDown = dIp[0];
                    }

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


            }

            if (l.Contains("CPU"))
            {
                sysAuditResults.cCPU = "[" + l.ToString().Substring(26).Replace("CC", "C");              


                string[] cCPUScores = sysAuditResults.cCPU.Split('=');
                string cCPUScore1 = cCPUScores[1].Replace("]  - [Processor", "").Trim().Replace("]  - Core Processor] ]", "").Trim();
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
                    string[] model = sysAuditResults.cCPU.Split(' ');

                    int pmodel = 0;
                   

                    if (model[11] == "CPU") {
                        pmodel = getCPU_Score(model[12]);
                    }
                    else if (model[12].Contains('X'))
                    {
                        pmodel = getCPU_Score(model[12] + " " + model[13]);
                    }
                    else if (model[9].Trim() == "AMD")
                    {
                        pmodel = getCPU_Score(model[10]);
                    }
                    else
                    {
                        pmodel = getCPU_Score(model[11]);
                    }
                    
                    
                    int cCPU = Convert.ToInt32(pmodel);                  

                    if(pmodel == 0 ){

                        sysAuditResults.CPUaResult = "Fail";
                        sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", **Needs manual processing**";
                        sysAuditResults.needsManualProcessing = true;
                        sysAuditResults.aResult = "Pending";

                    }
                    else
                    {
                        if (cCPU > 1000)
                        {
                            sysAuditResults.CPUaResult = "Pass";
                            sysAuditResults.cCPU = sysAuditResults.cCPU.Replace("No score given", "5.0");
                        }
                        else
                        {
                            sysAuditResults.CPUaResult = "Fail";
                            sysAuditResults.aFailedReason = sysAuditResults.aFailedReason + ", CPU insufficient";
                        }
                    }

                   


                }
            }

            //[OS = Windows 8.1] - 
            //[CPU Score = 5.5]  - 
            //[Processor = "FULL PROCESSOR NAME"] -
            //[Total RAM = 5GB]  - 
            //[Available Space = 602GB]  - 
            //[Download Speed = 31 Mbps] - 
            //[Upload Speed = 4.99 Mbps]  

         

            sysAuditResults.aResultSummary = "[OS = " + sysAuditResults.cOS + "] - " + sysAuditResults.cCPU + " - " + sysAuditResults.cRAM + " - " + sysAuditResults.cHDD + " - [Download Speed = " + sysAuditResults.cInternetDown + " Mbps] - [Upload Speed = " + sysAuditResults.cInternetUp + " Mbps]";
            return sysAuditResults;

        }
        struct SysAuditResults
        {

            private bool _needsmanualprocessing;
            private bool _isManual;

            private string _resultSent;
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
                    _internetUp = tf[idx].ToString();

                    //if (tf[0] == "" | tf[0] == ":" | tf[0] == "=")
                    //    idx = 1;
                    //if(tf[idx].ToString().Length > 2) { 
                    //    _internetUp = tf[idx].ToString().Substring(0, 3).Replace(".", "").Replace(";", "");
                    //}
                    //if (tf[idx].ToString().Length == 2)
                    //{
                    //    _internetUp = tf[idx].ToString().Replace(".", "").Replace(";", "");
                    //}
                    //if (tf[idx].ToString().Length == 1) {
                    //    _internetUp = tf[idx].ToString();
                    //}
                    //if (tf[idx].ToString().Length > 2)
                    //{
                    //    _internetUp = tf[idx].ToString().Substring(0, 3).Replace(".", "").Replace(";", "");



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

                    _internetDown = tf[idx].ToString();

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

            public bool IsManual
            {
                get
                {
                    return _isManual;
                }
                set
                {

                    _isManual = value;
                }
            }


            public string resultSent
            {
                get
                {
                    return _resultSent;
                }
                set
                {

                    _resultSent = value;
                }
            }

        }
        public static void sendMail(string recipient, string attachmentFilename, string cadidateName)
        {
           
                SmtpClient mySmtpClient = new SmtpClient("secure.emailsrvr.com", 25);
                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential("systemaudit@statesidebpo.com", "G4xxsfsdf$$$-fsdfxG3ffsdafsdfa4!df!!2eX");
                mySmtpClient.Credentials = basicAuthenticationInfo;

                MailAddress from = new MailAddress("systemaudit@statesidebpo.com");
                MailAddress to = new MailAddress(recipient);
                MailAddress cc = new MailAddress("recruiters@statesidebpo.com");
           

                if (isTESTING)
                {
                     to = new MailAddress("brodriguez@statesidebpo.com");
                     cc = new MailAddress("brodriguez@statesidebpo.com");
                }
                else
                {
                     cc = new MailAddress("recruiters@statesidebpo.com");
                }       

                MailMessage myMail = new MailMessage(from, to);
                myMail.IsBodyHtml = true;
                myMail.Subject = "System audit results for " + cadidateName;
                myMail.CC.Add(cc);

                if (isTESTING)
                {
                    myMail.Subject = "TESTING - System audit results for " + cadidateName;
                }

                string body = @"<p style =""font-size:21px"">Dear " + cadidateName + ",<br><br>" + "This email is to inform you of your system audit results. Please see the attachment. If you have any technical questions regarding your results, please reach out to us via email at <a mailto:winaudit@statesidebpo.com>winaudit@statesidebpo.com</a>.</p>";
                body = body + @"<p style =""font-size:18px""><i>(Please note: If you are unable to view the attachment, you may need to download and install Adobe Acrobat Reader DC or a similar program that allows the viewing of PDF documents)</i>";

                Attachment inlineStatetSideLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\StatesideLogo.png");
                Attachment inlineBitLeverLogo = new Attachment(@"\\filesvr4\IT\WinAudit\SysAudit App\Bit-LeverLogo.png");
                Attachment SysAuditResults = new Attachment(attachmentFilename);


                if (!isTESTING)
                {

                    SysAuditResults = new Attachment(@"\\Filesvr4\IT\WinAudit\Results_Archive\" + cadidateName + " SystemAudit Results.pdf");
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

            try
            {
               mySmtpClient.Send(myMail);
            }
             

            catch (System.Exception err)
            {
                throw new CustomException("Sending Individual Emails Results Error: ", err);
            }


        }
        private static void sendCompletionNotification(List<SysAuditResults> CandidatesList)
        {
           
            
                SmtpClient mySmtpClient = new SmtpClient("secure.emailsrvr.com", 25);
                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new
                System.Net.NetworkCredential("systemaudit@statesidebpo.com", "G4xxsfsdf$$$-fsdfxG3ffsdafsdfa4!df!!2eX");
                mySmtpClient.Credentials = basicAuthenticationInfo;

                // add from,to mailaddresses
                MailAddress to = new MailAddress("helpdesk@statesidebpo.com");
                MailAddress from = new MailAddress("systemaudit@statesidebpo.com");
                MailAddress cc = new MailAddress("recruiters@statesidebpo.com");


                if (isTESTING)
                {
                   to = new MailAddress("brodriguez@statesidebpo.com");
                   cc = new MailAddress("brodriguez@statesidebpo.com");
                }


                
                MailMessage myMail = new MailMessage(from, to);
                myMail.Subject = DateTime.Now + " SystemAudit Processing run completed successfully";
                myMail.CC.Add(cc);

                if (isTESTING)
                {
                    myMail.Subject = "TESTING - " + DateTime.Now + " SystemAudit Processing run completed successfully";

                }

                myMail.IsBodyHtml = true;


                string bd = "";

                if (CandidatesList.Count() == 1)
                {
                    bd = "<h2> " + CandidatesList.Count() + " Audit was processed</h2>";
                }
                else
                {
                    bd = "<h2> " + CandidatesList.Count() + " Audits were processed</h2>";
                }

                if (CandidatesList.Count() > 0)
                {

                    bd = bd + "<table border=" + "1" + " cellpadding=" + "10" + " width=" + "90%" + "><tbody>" +
                               "<tr><th bgcolor = " + "'#f49542'" + " style =" + "'padding: 5px 5px 5px 5px; color: white'" + " > Name </ th >" +
                               "<th  bgcolor =" + "'#f49542'" + " style = " + "'padding: 5px 0px 5px 0px;color: white'" + "> Email </ th >" +
                               "<th  bgcolor =" + "'#f49542'" + " style = " + "'padding: 5px 0px 5px 0px;color: white'" + "> Status </ th ></ tr >";

                    foreach (SysAuditResults c in CandidatesList)
                    {
                        string reason = c.aFailedReason;

                        if (c.aFailedReason != null)
                        {
                            reason = " - " + c.aFailedReason;
                        }

                        if (c.needsManualProcessing)
                        {
                            bd = bd + "<tr><td align=" + "'center'" + "width=" + "'23%'" + ">" + c.cName + "</td><td align=" + "'center'" + "width=" + "'33%'" + ">" + c.cEmail + "</td><td align=" + "'center'" + " width=" + "'43%'" + ">" + c.aResult + reason + "</td></tr>";


                        }
                        else
                        {
                            if (c.aResult == "Fail")
                            {
                                bd = bd + "<tr><td align=" + "'center'" + "width =" + "'23%'" + ">" + c.cName + "</td><td align=" + "'center'" + "width=" + "'43%'" + ">" + c.cEmail + "</td><td  align=" + "'center'" + "width=" + "'43%'" + "><font color='red'>" + c.aResult + reason + "</font></td></tr>";
                            }
                            if (c.aResult == "Pass")
                            {
                                bd = bd + "<tr><td  align=" + "'center'" + "width=" + "'23%'" + ">" + c.cName + "</td><td align=" + "'center'" + "width=" + "'28%'" + ">" + c.cEmail + "</td><td ' align=" + "'center'" + "width=" + "'43%'" + "><font color='green'>" + c.aResult + reason + "</font></td></tr>";
                            }

                            if (c.aResult == "Pending")
                            {
                                bd = bd + "<tr><td align=" + "'center'" + "  width=" + "'23%'" + ">" + c.cName + "</td><td align=" + "'center'" + "width=" + "'28%'" + ">" + c.cEmail + "</td><td  align=" + "'center'" + "width=" + "'43%'" + "><font color='blue'>" + c.aResult + reason + "</font></td></tr>";
                            }

                            sendMail(c.cEmail, c.attachmentFilename, c.cName);
                        }


                    }

                    bd = bd + "</tbody></table>";

                }




                myMail.Body = Regex.Replace(bd, @"[^\u0000-\u007F]", " ");
              

            try
            {
                mySmtpClient.Send(myMail);
            }

            catch (System.Exception err)
            {
                throw new CustomException("Sending Notification Email Error: ", err);
            }
        }

        private static int getCPU_Score(string model)
        {


            var dict = new Dictionary<string, Int32>();
            dict.Add("686 Gen", 288);
            dict.Add("AMD A4 Micro-6400T APU", 1544);
            dict.Add("AMD A4 PRO-3340B", 2466);
            dict.Add("AMD A4 PRO-7300B APU", 2342);
            dict.Add("AMD A4 PRO-7350B", 2794);
            dict.Add("AMD A4-1200 APU", 632);
            dict.Add("AMD A4-1250 APU", 610);
            dict.Add("AMD A4-3300 APU", 1505);
            dict.Add("AMD A4-3300M APU", 1194);
            dict.Add("AMD A4-3305M APU", 1218);
            dict.Add("AMD A4-3310MX APU", 1310);
            dict.Add("AMD A4-3320M APU", 1247);
            dict.Add("AMD A4-3330MX APU", 1217);
            dict.Add("AMD A4-3400 APU", 1595);
            dict.Add("AMD A4-3420 APU", 1679);
            dict.Add("AMD A4-4000 APU", 1816);
            dict.Add("AMD A4-4020 APU", 1822);
            dict.Add("AMD A4-4300M APU", 1679);
            dict.Add("AMD A4-4355M APU", 1161);
            dict.Add("AMD A4-5000 APU", 1900);
            dict.Add("AMD A4-5050 APU", 2251);
            dict.Add("AMD A4-5100 APU", 2067);
            dict.Add("AMD A4-5150M APU", 1860);
            dict.Add("AMD A4-5300 APU", 1998);
            dict.Add("AMD A4-5300B APU", 2039);
            dict.Add("AMD A4-6210 APU", 2134);
            dict.Add("AMD A4-6250J APU", 2552);
            dict.Add("AMD A4-6300 APU", 2224);
            dict.Add("AMD A4-6300B APU", 2042);
            dict.Add("AMD A4-6320 APU", 2310);
            dict.Add("AMD A4-7210 APU", 2511);
            dict.Add("AMD A4-7300 APU", 2255);
            dict.Add("AMD A6 Micro-6500T APU", 1834);
            dict.Add("AMD A6 PRO-7050B APU", 1573);
            dict.Add("AMD A6 PRO-7400B", 2697);
            dict.Add("AMD A6-1450 APU", 1526);
            dict.Add("AMD A6-3400M APU", 1902);
            dict.Add("AMD A6-3410MX APU", 2075);
            dict.Add("AMD A6-3420M APU", 2022);
            dict.Add("AMD A6-3430MX APU", 2224);
            dict.Add("AMD A6-3500 APU", 2002);
            dict.Add("AMD A6-3600 APU", 2713);
            dict.Add("AMD A6-3620 APU", 2823);
            dict.Add("AMD A6-3650 APU", 3175);
            dict.Add("AMD A6-3670 APU", 3209);
            dict.Add("AMD A6-4400M APU", 1621);
            dict.Add("AMD A6-4455M APU", 1279);
            dict.Add("AMD A6-5200 APU", 2400);
            dict.Add("AMD A6-5345M APU", 1481);
            dict.Add("AMD A6-5350M APU", 1902);
            dict.Add("AMD A6-5357M APU", 1867);
            dict.Add("AMD A6-5400B APU", 2225);
            dict.Add("AMD A6-5400K APU", 2139);
            dict.Add("AMD A6-6310 APU", 2406);
            dict.Add("AMD A6-6400B APU", 2414);
            dict.Add("AMD A6-6400K APU", 2290);
            dict.Add("AMD A6-6420B APU", 2612);
            dict.Add("AMD A6-6420K APU", 2256);
            dict.Add("AMD A6-7000", 1681);
            dict.Add("AMD A6-7310 APU", 2650);
            dict.Add("AMD A6-7400K APU", 2796);
            dict.Add("AMD A6-7470K", 2778);
            dict.Add("AMD A6-8500P", 2183);
            dict.Add("AMD A6-8550", 3104);
            dict.Add("AMD A6-9210", 1924);
            dict.Add("AMD A6-9500", 3212);
            dict.Add("AMD A8 PRO-7150B APU", 2732);
            dict.Add("AMD A8 PRO-7600B APU", 4823);
            dict.Add("AMD A8-3500M APU", 1974);
            dict.Add("AMD A8-3510MX APU", 2389);
            dict.Add("AMD A8-3520M APU", 2202);
            dict.Add("AMD A8-3530MX APU", 2380);
            dict.Add("AMD A8-3550MX APU", 2629);
            dict.Add("AMD A8-3800 APU", 3096);
            dict.Add("AMD A8-3820 APU", 3151);
            dict.Add("AMD A8-3850 APU", 3489);
            dict.Add("AMD A8-3870K APU", 3597);
            dict.Add("AMD A8-4500M APU", 2655);
            dict.Add("AMD A8-4555M APU", 2119);
            dict.Add("AMD A8-5500 APU", 3985);
            dict.Add("AMD A8-5500B APU", 3821);
            dict.Add("AMD A8-5545M APU", 2553);
            dict.Add("AMD A8-5550M APU", 3002);
            dict.Add("AMD A8-5557M APU", 2962);
            dict.Add("AMD A8-5600K APU", 4324);
            dict.Add("AMD A8-6410 APU", 2531);
            dict.Add("AMD A8-6500 APU", 4386);
            dict.Add("AMD A8-6500B APU", 4582);
            dict.Add("AMD A8-6500T APU", 3221);
            dict.Add("AMD A8-6600K APU", 4570);
            dict.Add("AMD A8-7050", 1986);
            dict.Add("AMD A8-7100 APU", 2827);
            dict.Add("AMD A8-7200P", 3366);
            dict.Add("AMD A8-7410 APU", 2700);
            dict.Add("AMD A8-7600 APU", 5138);
            dict.Add("AMD A8-7650K", 4994);
            dict.Add("AMD A8-7670K", 5072);
            dict.Add("AMD A8-8600P", 3251);
            dict.Add("AMD A8-8650", 5488);
            dict.Add("AMD A8-9600", 5270);
            dict.Add("AMD A9-9400", 2237);
            dict.Add("AMD A9-9410", 2598);
            dict.Add("AMD A10 Micro-6700T APU", 1787);
            dict.Add("AMD A10 PRO-7350B APU", 3274);
            dict.Add("AMD A10 PRO-7800B APU", 5424);
            dict.Add("AMD A10 PRO-7850B APU", 5614);
            dict.Add("AMD A10-4600M APU", 3111);
            dict.Add("AMD A10-4655M APU", 2581);
            dict.Add("AMD A10-4657M APU", 2906);
            dict.Add("AMD A10-5700 APU", 4201);
            dict.Add("AMD A10-5745M APU", 2765);
            dict.Add("AMD A10-5750M APU", 3343);
            dict.Add("AMD A10-5757M APU", 3099);
            dict.Add("AMD A10-5800B APU", 4572);
            dict.Add("AMD A10-5800K APU", 4631);
            dict.Add("AMD A10-6700 APU", 4598);
            dict.Add("AMD A10-6700T APU", 3702);
            dict.Add("AMD A10-6790K APU", 4665);
            dict.Add("AMD A10-6800B APU", 4990);
            dict.Add("AMD A10-6800K APU", 4894);
            dict.Add("AMD A10-7300 APU", 2938);
            dict.Add("AMD A10-7400P", 3664);
            dict.Add("AMD A10-7700K APU", 5163);
            dict.Add("AMD A10-7800 APU", 5070);
            dict.Add("AMD A10-7850K APU", 5503);
            dict.Add("AMD A10-7860K", 5128);
            dict.Add("AMD A10-7870K", 5497);
            dict.Add("AMD A10-7890K", 5746);
            dict.Add("AMD A10-8700P", 3461);
            dict.Add("AMD A10-8750", 5177);
            dict.Add("AMD A10-8850", 5613);
            dict.Add("AMD A10-9600P", 3872);
            dict.Add("AMD A10-9630P", 4562);
            dict.Add("AMD A10-9700", 5486);
            dict.Add("AMD A12-9700P", 3887);
            dict.Add("AMD A12-9800", 5737);
            dict.Add("AMD Athlon64 X2 Dual Core 3800 +", 997);
            dict.Add("AMD Athlon64 X2 Dual Core 4200 +", 1006);
            dict.Add("AMD Athlon64 X2 Dual Core 4600 +", 1108);
            dict.Add("AMD Athlon 4", 279);
            dict.Add("AMD Athlon 64 2000 +", 116);
            dict.Add("AMD Athlon 64 2200 +", 575);
            dict.Add("AMD Athlon 64 2600 +", 374);
            dict.Add("AMD Athlon 64 2800 +", 437);
            dict.Add("AMD Athlon 64 3000 +", 461);
            dict.Add("AMD Athlon 64 3200 +", 491);
            dict.Add("AMD Athlon 64 3300 +", 576);
            dict.Add("AMD Athlon 64 3400 +", 544);
            dict.Add("AMD Athlon 64 3500 +", 537);
            dict.Add("AMD Athlon 64 3600 +", 601);
            dict.Add("AMD Athlon 64 3700 +", 599);
            dict.Add("AMD Athlon 64 3800 +", 588);
            dict.Add("AMD Athlon 64 4000 +", 606);
            dict.Add("AMD Athlon 64 FX-25", 794);
            dict.Add("AMD Athlon 64 FX-34", 380);
            dict.Add("AMD Athlon 64 FX-51", 437);
            dict.Add("AMD Athlon 64 FX-53", 646);
            dict.Add("AMD Athlon 64 FX-55", 721);
            dict.Add("AMD Athlon 64 FX-57", 730);
            dict.Add("AMD Athlon 64 FX-59", 715);
            dict.Add("AMD Athlon 64 FX-60 Dual Core", 1287);
            dict.Add("AMD Athlon 64 FX-62 Dual Core", 1571);
            dict.Add("AMD Athlon 64 FX-72", 1794);
            dict.Add("AMD Athlon 64 FX-74", 1495);
            dict.Add("AMD Athlon 64 X2 3800 +", 958);
            dict.Add("AMD Athlon 64 X2 4200 +", 1285);
            dict.Add("AMD Athlon 64 X2 4600 +", 1365);
            dict.Add("AMD Athlon 64 X2 Dual Core 3400 +", 1003);
            dict.Add("AMD Athlon 64 X2 Dual Core 3600 +", 970);
            dict.Add("AMD Athlon 64 X2 Dual Core 3800 +", 1003);
            dict.Add("AMD Athlon 64 X2 Dual Core 4000 +", 1041);
            dict.Add("AMD Athlon 64 X2 Dual Core 4200 +", 1099);
            dict.Add("AMD Athlon 64 X2 Dual Core 4400 +", 1148);
            dict.Add("AMD Athlon 64 X2 Dual Core 4600 +", 1220);
            dict.Add("AMD Athlon 64 X2 Dual Core 4800 +", 1264);
            dict.Add("AMD Athlon 64 X2 Dual Core 5000 +", 1311);
            dict.Add("AMD Athlon 64 X2 Dual Core 5200 +", 1384);
            dict.Add("AMD Athlon 64 X2 Dual Core 5400 +", 1449);
            dict.Add("AMD Athlon 64 X2 Dual Core 5600 +", 1475);
            dict.Add("AMD Athlon 64 X2 Dual Core 5800 +", 1536);
            dict.Add("AMD Athlon 64 X2 Dual Core 6000 +", 1604);
            dict.Add("AMD Athlon 64 X2 Dual Core 6400 +", 1792);
            dict.Add("AMD Athlon 64 X2 Dual Core BE-230", 1095);
            dict.Add("AMD Athlon 64 X2 Dual Core BE-2300", 1082);
            dict.Add("AMD Athlon 64 X2 Dual Core BE-2350", 1044);
            dict.Add("AMD Athlon 64 X2 Dual Core TK-53", 839);
            dict.Add("AMD Athlon 64 X2 Dual Core TK-55", 839);
            dict.Add("AMD Athlon 64 X2 Dual Core TK-57", 971);
            dict.Add("AMD Athlon 64 X2 Dual-Core TK-42", 876);
            dict.Add("AMD Athlon 64 X2 Dual-Core TK-53", 804);
            dict.Add("AMD Athlon 64 X2 Dual-Core TK-55", 881);
            dict.Add("AMD Athlon 64 X2 Dual-Core TK-57", 918);
            dict.Add("AMD Athlon 64 X2 QL-60", 923);
            dict.Add("AMD Athlon 64 X2 QL-62", 1026);
            dict.Add("AMD Athlon 64 X2 QL-64", 974);
            dict.Add("AMD Athlon 64 X2 QL-65", 1035);
            dict.Add("AMD Athlon 64 X2 QL-66", 939);
            dict.Add("AMD Athlon 64 X2 QL-67", 1006);
            dict.Add("AMD Athlon 64 X2 TK-55", 955);
            dict.Add("AMD Athlon 64 X2 TK-57", 988);
            dict.Add("AMD Athlon 1500 +", 294);
            dict.Add("AMD Athlon 1640B", 668);
            dict.Add("AMD Athlon 1700 +", 349);
            dict.Add("AMD Athlon 2000 +", 390);
            dict.Add("AMD Athlon 2100 +", 412);
            dict.Add("AMD Athlon 2200 +", 363);
            dict.Add("AMD Athlon 2400 +", 462);
            dict.Add("AMD Athlon 2500 +", 374);
            dict.Add("AMD Athlon 2600 +", 444);
            dict.Add("AMD Athlon 2650e", 407);
            dict.Add("AMD Athlon 2800 +", 447);
            dict.Add("AMD Athlon 2850e", 498);
            dict.Add("AMD Athlon 3100 +", 485);
            dict.Add("AMD Athlon 3200 +", 458);
            dict.Add("AMD Athlon 5000 Dual-Core", 1285);
            dict.Add("AMD Athlon 5150 APU", 2091);
            dict.Add("AMD Athlon 5200 Dual-Core", 1238);
            dict.Add("AMD Athlon 5350 APU", 2577);
            dict.Add("AMD Athlon 5370 APU", 2864);
            dict.Add("AMD Athlon 7450 Dual-Core", 1421);
            dict.Add("AMD Athlon 7550 Dual-Core", 1453);
            dict.Add("AMD Athlon 7750 Dual-Core", 1577);
            dict.Add("AMD Athlon 7850 Dual-Core", 1717);
            dict.Add("AMD Athlon 7850 Dual-Core 2.8G", 288);
            dict.Add("AMD Athlon Dual Core 4050e", 1086);
            dict.Add("AMD Athlon Dual Core 4450B", 1170);
            dict.Add("AMD Athlon Dual Core 4450e", 1101);
            dict.Add("AMD Athlon Dual Core 4850B", 1471);
            dict.Add("AMD Athlon Dual Core 4850e", 1314);
            dict.Add("AMD Athlon Dual Core 5000B", 1298);
            dict.Add("AMD Athlon Dual Core 5050e", 1347);
            dict.Add("AMD Athlon Dual Core 5200B", 1334);
            dict.Add("AMD Athlon Dual Core 5400B", 1425);
            dict.Add("AMD Athlon Dual Core 5600B", 1304);
            dict.Add("AMD Athlon II 160u", 552);
            dict.Add("AMD Athlon II 170u", 624);
            dict.Add("AMD Athlon II Dual-Core M300", 1118);
            dict.Add("AMD Athlon II Dual-Core M320", 1147);
            dict.Add("AMD Athlon II Dual-Core M340", 1040);
            dict.Add("AMD Athlon II N330 Dual-Core", 1154);
            dict.Add("AMD Athlon II N350 Dual-Core", 1225);
            dict.Add("AMD Athlon II N370 Dual-Core", 1559);
            dict.Add("AMD Athlon II Neo K125", 474);
            dict.Add("AMD Athlon II Neo K145", 543);
            dict.Add("AMD Athlon II Neo K325 Dual-Core", 785);
            dict.Add("AMD Athlon II Neo K345 Dual-Core", 856);
            dict.Add("AMD Athlon II Neo N36L Dual-Core", 817);
            dict.Add("AMD Athlon II P320 Dual-Core", 1214);
            dict.Add("AMD Athlon II P340 Dual-Core", 1253);
            dict.Add("AMD Athlon II P360 Dual-Core", 1318);
            dict.Add("AMD Athlon II X2 210e", 1486);
            dict.Add("AMD Athlon II X2 215", 1581);
            dict.Add("AMD Athlon II X2 220", 1631);
            dict.Add("AMD Athlon II X2 225", 1526);
            dict.Add("AMD Athlon II X2 235e", 1658);
            dict.Add("AMD Athlon II X2 240", 1641);
            dict.Add("AMD Athlon II X2 240e", 1699);
            dict.Add("AMD Athlon II X2 245", 1688);
            dict.Add("AMD Athlon II X2 245e", 1718);
            dict.Add("AMD Athlon II X2 250", 1749);
            dict.Add("AMD Athlon II X2 250e", 1949);
            dict.Add("AMD Athlon II X2 250u", 1005);
            dict.Add("AMD Athlon II X2 255", 1843);
            dict.Add("AMD Athlon II X2 260", 1887);
            dict.Add("AMD Athlon II X2 260u", 1119);
            dict.Add("AMD Athlon II X2 265", 1870);
            dict.Add("AMD Athlon II X2 270", 1988);
            dict.Add("AMD Athlon II X2 270u", 1350);
            dict.Add("AMD Athlon II X2 280", 2183);
            dict.Add("AMD Athlon II X2 4300e", 1464);
            dict.Add("AMD Athlon II X2 4400e", 1618);
            dict.Add("AMD Athlon II X2 4450e", 1504);
            dict.Add("AMD Athlon II X2 B22", 1699);
            dict.Add("AMD Athlon II X2 B24", 1798);
            dict.Add("AMD Athlon II X2 B26", 1804);
            dict.Add("AMD Athlon II X2 B28", 2118);
            dict.Add("AMD Athlon II X3 400e", 2018);
            dict.Add("AMD Athlon II X3 405e", 2041);
            dict.Add("AMD Athlon II X3 415e", 2246);
            dict.Add("AMD Athlon II X3 420e", 2286);
            dict.Add("AMD Athlon II X3 425", 2325);
            dict.Add("AMD Athlon II X3 425e", 2239);
            dict.Add("AMD Athlon II X3 435", 2475);
            dict.Add("AMD Athlon II X3 440", 2568);
            dict.Add("AMD Athlon II X3 445", 2586);
            dict.Add("AMD Athlon II X3 450", 2627);
            dict.Add("AMD Athlon II X3 455", 2779);
            dict.Add("AMD Athlon II X3 460", 2876);
            dict.Add("AMD Athlon II X4 553", 3814);
            dict.Add("AMD Athlon II X4 555", 3830);
            dict.Add("AMD Athlon II X4 559", 4157);
            dict.Add("AMD Athlon II X4 600e", 2342);
            dict.Add("AMD Athlon II X4 605e", 2739);
            dict.Add("AMD Athlon II X4 610e", 2811);
            dict.Add("AMD Athlon II X4 615e", 2877);
            dict.Add("AMD Athlon II X4 620", 2950);
            dict.Add("AMD Athlon II X4 620e", 2906);
            dict.Add("AMD Athlon II X4 630", 3147);
            dict.Add("AMD Athlon II X4 631 Quad-Core", 3152);
            dict.Add("AMD Athlon II X4 635", 3269);
            dict.Add("AMD Athlon II X4 638 Quad-Core", 3198);
            dict.Add("AMD Athlon II X4 640", 3328);
            dict.Add("AMD Athlon II X4 641 Quad-Core", 3399);
            dict.Add("AMD Athlon II X4 645", 3496);
            dict.Add("AMD Athlon II X4 650", 3469);
            dict.Add("AMD Athlon II X4 651 Quad-Core", 3569);
            dict.Add("AMD Athlon II X4 655", 2880);
            dict.Add("AMD Athlon II X4 6400e", 2594);
            dict.Add("AMD Athlon L110", 306);
            dict.Add("AMD Athlon LE-1600", 561);
            dict.Add("AMD Athlon LE-1620", 571);
            dict.Add("AMD Athlon LE-1640", 623);
            dict.Add("AMD Athlon LE-1660", 682);
            dict.Add("AMD Athlon MP", 230);
            dict.Add("AMD Athlon MP 1700 +", 339);
            dict.Add("AMD Athlon MP 2400 +", 436);
            dict.Add("AMD Athlon MP 2800 +", 497);
            dict.Add("AMD Athlon Neo MV-40", 406);
            dict.Add("AMD Athlon Neo X2 Dual Core 6850e", 1002);
            dict.Add("AMD Athlon Neo X2 Dual Core L325", 770);
            dict.Add("AMD Athlon Neo X2 Dual Core L335", 814);
            dict.Add("AMD Athlon QI-46", 473);
            dict.Add("AMD Athlon TF-20", 417);
            dict.Add("AMD Athlon TF-36", 329);
            dict.Add("AMD Athlon X2 215", 1575);
            dict.Add("AMD Athlon X2 235e", 1806);
            dict.Add("AMD Athlon X2 240", 1575);
            dict.Add("AMD Athlon X2 240e", 1670);
            dict.Add("AMD Athlon X2 250", 1674);
            dict.Add("AMD Athlon X2 255", 1627);
            dict.Add("AMD Athlon X2 280", 2303);
            dict.Add("AMD Athlon X2 340 Dual Core", 1889);
            dict.Add("AMD Athlon X2 370K Dual Core", 2229);
            dict.Add("AMD Athlon X2 440", 1424);
            dict.Add("AMD Athlon X2 Dual Core 3250e", 774);
            dict.Add("AMD Athlon X2 Dual Core 6850e", 998);
            dict.Add("AMD Athlon X2 Dual Core BE-2300", 963);
            dict.Add("AMD Athlon X2 Dual Core BE-2350", 1098);
            dict.Add("AMD Athlon X2 Dual Core BE-2400", 1236);
            dict.Add("AMD Athlon X2 Dual Core BE-2450", 1251);
            dict.Add("AMD Athlon X2 Dual Core L310", 644);
            dict.Add("AMD Athlon X2 Dual Core LS-5800", 1009);
            dict.Add("AMD Athlon X3 425", 2191);
            dict.Add("AMD Athlon X3 435", 2574);
            dict.Add("AMD Athlon X3 440", 2078);
            dict.Add("AMD Athlon X3 445", 2314);
            dict.Add("AMD Athlon X3 450", 3064);
            dict.Add("AMD Athlon X3 455", 2401);
            dict.Add("AMD Athlon X4 605e", 2505);
            dict.Add("AMD Athlon X4 620", 2855);
            dict.Add("AMD Athlon X4 635", 2891);
            dict.Add("AMD Athlon X4 640", 3172);
            dict.Add("AMD Athlon X4 645", 3595);
            dict.Add("AMD Athlon X4 740 Quad Core", 3992);
            dict.Add("AMD Athlon X4 750 Quad Core", 4551);
            dict.Add("AMD Athlon X4 750K Quad Core", 4224);
            dict.Add("AMD Athlon X4 760K Quad Core", 4577);
            dict.Add("AMD Athlon X4 840", 5013);
            dict.Add("AMD Athlon X4 845", 5434);
            dict.Add("AMD Athlon X4 860K", 5477);
            dict.Add("AMD Athlon X4 870K", 5282);
            dict.Add("AMD Athlon X4 880K", 5571);
            dict.Add("AMD Athlon XP1600 +", 281);
            dict.Add("AMD Athlon XP1700 +", 234);
            dict.Add("AMD Athlon XP1900 +", 353);
            dict.Add("AMD Athlon XP2100 +", 287);
            dict.Add("AMD Athlon XP2200 +", 401);
            dict.Add("AMD Athlon XP2400 +", 391);
            dict.Add("AMD Athlon XP 1500", 282);
            dict.Add("AMD Athlon XP 1500 +", 257);
            dict.Add("AMD Athlon XP 1600 +", 266);
            dict.Add("AMD Athlon XP 1700", 344);
            dict.Add("AMD Athlon XP 1700 +", 268);
            dict.Add("AMD Athlon XP 1800 +", 300);
            dict.Add("AMD Athlon XP 1900 +", 305);
            dict.Add("AMD Athlon XP 2000 +", 300);
            dict.Add("AMD Athlon XP 2100 +", 323);
            dict.Add("AMD Athlon XP 2200 +", 327);
            dict.Add("AMD Athlon XP 2300 +", 464);
            dict.Add("AMD Athlon XP 2400 +", 362);
            dict.Add("AMD Athlon XP 2500 +", 355);
            dict.Add("AMD Athlon XP 2600 +", 372);
            dict.Add("AMD Athlon XP 2700 +", 393);
            dict.Add("AMD Athlon XP 2800 +", 398);
            dict.Add("AMD Athlon XP 2900 +", 459);
            dict.Add("AMD Athlon XP 3000 +", 402);
            dict.Add("AMD Athlon XP 3100 +", 433);
            dict.Add("AMD Athlon XP 3200 +", 431);
            dict.Add("AMD Athlon XP 3400 +", 570);
            dict.Add("AMD Athlon XP Pro3 +", 367);
            dict.Add("AMD Athlon XP-M", 541);
            dict.Add("AMD C-30", 310);
            dict.Add("AMD C-50", 462);
            dict.Add("AMD C-60", 448);
            dict.Add("AMD C-60 APU", 543);
            dict.Add("AMD C-70 APU", 577);
            dict.Add("AMD E1 Micro-6200T APU", 960);
            dict.Add("AMD E1-1200 APU", 666);
            dict.Add("AMD E1-1500 APU", 693);
            dict.Add("AMD E1-2100 APU", 625);
            dict.Add("AMD E1-2200 APU", 767);
            dict.Add("AMD E1-2500 APU", 880);
            dict.Add("AMD E1-6010 APU", 844);
            dict.Add("AMD E1-6015 APU", 880);
            dict.Add("AMD E1-7010 APU", 974);
            dict.Add("AMD E2-1800 APU", 823);
            dict.Add("AMD E2-2000 APU", 829);
            dict.Add("AMD E2-3000 APU", 1080);
            dict.Add("AMD E2-3000M APU", 1095);
            dict.Add("AMD E2-3200 APU", 1454);
            dict.Add("AMD E2-3800 APU", 1619);
            dict.Add("AMD E2-6110 APU", 1890);
            dict.Add("AMD E2-7110 APU", 2261);
            dict.Add("AMD E2-9010", 1921);
            dict.Add("AMD E-240", 317);
            dict.Add("AMD E-300 APU", 616);
            dict.Add("AMD E-350", 748);
            dict.Add("AMD E-350 APU", 774);
            dict.Add("AMD E-350D APU", 717);
            dict.Add("AMD E-450 APU", 769);
            dict.Add("AMD Embedded R-Series RX-421BD", 4858);
            dict.Add("AMD FirePro A320 APU", 4713);
            dict.Add("AMD FX-670K Quad-Core", 4520);
            dict.Add("AMD FX-770K Quad-Core", 4907);
            dict.Add("AMD FX-870K Quad Core", 5226);
            dict.Add("AMD FX-4100 Quad-Core", 4047);
            dict.Add("AMD FX-4130 Quad-Core", 4157);
            dict.Add("AMD FX-4150 Quad-Core", 4576);
            dict.Add("AMD FX-4170 Quad-Core", 4837);
            dict.Add("AMD FX-4200 Quad-Core", 4264);
            dict.Add("AMD FX-4300 Quad-Core", 4644);
            dict.Add("AMD FX-4320", 5010);
            dict.Add("AMD FX-4330", 5192);
            dict.Add("AMD FX-4350 Quad-Core", 5296);
            dict.Add("AMD FX-6100 Six-Core", 5406);
            dict.Add("AMD FX-6120 Six-Core", 5746);
            dict.Add("AMD FX-6130 Six-Core", 6215);
            dict.Add("AMD FX-6200 Six-Core", 6111);
            dict.Add("AMD FX-6300 Six-Core", 6353);
            dict.Add("AMD FX-6330 Six-Core", 6437);
            dict.Add("AMD FX-6350 Six-Core", 6965);
            dict.Add("AMD FX-7500 APU", 3208);
            dict.Add("AMD FX-7600P", 4053);
            dict.Add("AMD FX-7600P APU", 4627);
            dict.Add("AMD FX-8100 Eight-Core", 6080);
            dict.Add("AMD FX-8120 Eight-Core", 6582);
            dict.Add("AMD FX-8140 Eight-Core", 5854);
            dict.Add("AMD FX-8150 Eight-Core", 7624);
            dict.Add("AMD FX-8300 Eight-Core", 7667);
            dict.Add("AMD FX-8310 Eight-Core", 7753);
            dict.Add("AMD FX-8320 Eight-Core", 8018);
            dict.Add("AMD FX-8320E Eight-Core", 7502);
            dict.Add("AMD FX-8350 Eight-Core", 8943);
            dict.Add("AMD FX-8370 Eight-Core", 8989);
            dict.Add("AMD FX-8370E Eight-Core", 7787);
            dict.Add("AMD FX-8800P", 4184);
            dict.Add("AMD FX-9370 Eight-Core", 9525);
            dict.Add("AMD FX-9590 Eight-Core", 10327);
            dict.Add("AMD FX-9800P", 4106);
            dict.Add("AMD FX-9830P", 5053);
            dict.Add("AMD FX-B4150 Quad-Core", 4611);
            dict.Add("AMD G-T40E", 485);
            dict.Add("AMD G-T40N", 516);
            dict.Add("AMD G-T44R", 227);
            dict.Add("AMD G-T48E", 716);
            dict.Add("AMD G-T52R", 303);
            dict.Add("AMD G-T56E", 751);
            dict.Add("AMD G-T56N", 793);
            dict.Add("AMD Geode NX", 264);
            dict.Add("AMD Geode NX 1750", 291);
            dict.Add("AMD GX-210JA SOC", 401);
            dict.Add("AMD GX-212JC SOC", 744);
            dict.Add("AMD GX-217GA SOC", 1080);
            dict.Add("AMD GX-218GL SOC", 1115);
            dict.Add("AMD GX-222GC SOC", 1180);
            dict.Add("AMD GX-412HC", 1462);
            dict.Add("AMD GX-415GA SOC", 1940);
            dict.Add("AMD GX-420CA SOC", 2299);
            dict.Add("AMD GX-424CC SOC", 2375);
            dict.Add("AMD K7", 403);
            dict.Add("AMD Opteron 140", 396);
            dict.Add("AMD Opteron 142", 428);
            dict.Add("AMD Opteron 144", 511);
            dict.Add("AMD Opteron 146", 532);
            dict.Add("AMD Opteron 148", 603);
            dict.Add("AMD Opteron 150", 604);
            dict.Add("AMD Opteron 152", 658);
            dict.Add("AMD Opteron 154", 703);
            dict.Add("AMD Opteron 165", 899);
            dict.Add("AMD Opteron 170", 1097);
            dict.Add("AMD Opteron 175", 1167);
            dict.Add("AMD Opteron 180", 1232);
            dict.Add("AMD Opteron 185", 1359);
            dict.Add("AMD Opteron 246", 632);
            dict.Add("AMD Opteron 248", 476);
            dict.Add("AMD Opteron 250", 649);
            dict.Add("AMD Opteron 252", 760);
            dict.Add("AMD Opteron 254", 693);
            dict.Add("AMD Opteron 256", 963);
            dict.Add("AMD Opteron 265", 963);
            dict.Add("AMD Opteron 270", 1083);
            dict.Add("AMD Opteron 275", 1291);
            dict.Add("AMD Opteron 280", 1272);
            dict.Add("AMD Opteron 285", 1470);
            dict.Add("AMD Opteron 290", 1464);
            dict.Add("AMD Opteron 1210", 923);
            dict.Add("AMD Opteron 1210 HE", 971);
            dict.Add("AMD Opteron 1212", 1015);
            dict.Add("AMD Opteron 1212 HE", 1308);
            dict.Add("AMD Opteron 1214", 1191);
            dict.Add("AMD Opteron 1214 HE", 1299);
            dict.Add("AMD Opteron 1216", 1352);
            dict.Add("AMD Opteron 1216 HE", 1172);
            dict.Add("AMD Opteron 1218", 1468);
            dict.Add("AMD Opteron 1218 HE", 1444);
            dict.Add("AMD Opteron 1220", 1653);
            dict.Add("AMD Opteron 1222", 1757);
            dict.Add("AMD Opteron 1352", 2368);
            dict.Add("AMD Opteron 1354", 2609);
            dict.Add("AMD Opteron 1356", 2765);
            dict.Add("AMD Opteron 1381", 3065);
            dict.Add("AMD Opteron 2212", 1270);
            dict.Add("AMD Opteron 2214", 1358);
            dict.Add("AMD Opteron 2214 HE", 1415);
            dict.Add("AMD Opteron 2216", 1518);
            dict.Add("AMD Opteron 2218", 1594);
            dict.Add("AMD Opteron 2220", 1701);
            dict.Add("AMD Opteron 2220 SE", 1398);
            dict.Add("AMD Opteron 2224 SE", 1919);
            dict.Add("AMD Opteron 2350", 2820);
            dict.Add("AMD Opteron 2356", 2814);
            dict.Add("AMD Opteron 2374 HE", 2951);
            dict.Add("AMD Opteron 2376", 2283);
            dict.Add("AMD Opteron 2378", 3179);
            dict.Add("AMD Opteron 2380", 2640);
            dict.Add("AMD Opteron 2384", 3538);
            dict.Add("AMD Opteron 2386 SE", 3538);
            dict.Add("AMD Opteron 2393 SE", 2246);
            dict.Add("AMD Opteron 2427", 3069);
            dict.Add("AMD Opteron 2431", 4516);
            dict.Add("AMD Opteron 2435", 4708);
            dict.Add("AMD Opteron 3260 HE", 3259);
            dict.Add("AMD Opteron 3280", 5271);
            dict.Add("AMD Opteron 3320 EE", 2532);
            dict.Add("AMD Opteron 3350 HE", 4009);
            dict.Add("AMD Opteron 3365", 5803);
            dict.Add("AMD Opteron 3380", 6384);
            dict.Add("AMD Opteron 4122", 2942);
            dict.Add("AMD Opteron 4162 EE", 3166);
            dict.Add("AMD Opteron 4184", 4418);
            dict.Add("AMD Opteron 4280", 6430);
            dict.Add("AMD Opteron 4284", 6660);
            dict.Add("AMD Opteron 4332 HE", 5401);
            dict.Add("AMD Opteron 4365 EE", 5025);
            dict.Add("AMD Opteron 6128", 4742);
            dict.Add("AMD Opteron 6134", 2938);
            dict.Add("AMD Opteron 6136", 4631);
            dict.Add("AMD Opteron 6164 HE", 5351);
            dict.Add("AMD Opteron 6174", 6412);
            dict.Add("AMD Opteron 6212", 5776);
            dict.Add("AMD Opteron 6234", 5979);
            dict.Add("AMD Opteron 6238", 7316);
            dict.Add("AMD Opteron 6272", 6748);
            dict.Add("AMD Opteron 6274", 4865);
            dict.Add("AMD Opteron 6276", 9011);
            dict.Add("AMD Opteron 6282 SE", 9116);
            dict.Add("AMD Opteron 6287 SE", 9595);
            dict.Add("AMD Opteron 6328", 7813);
            dict.Add("AMD Opteron 6366 HE", 7503);
            dict.Add("AMD Opteron 6376", 9414);
            dict.Add("AMD Opteron 6378", 2152);
            dict.Add("AMD Opteron 6380", 10082);
            dict.Add("AMD Opteron 8354", 2255);
            dict.Add("AMD Opteron 8439 SE", 3813);
            dict.Add("AMD Phenom 2 X4 12000", 3005);
            dict.Add("AMD Phenom 7950 Quad-Core", 2991);
            dict.Add("AMD Phenom 8250 Triple-Core", 1599);
            dict.Add("AMD Phenom 8250e Triple-Core", 1602);
            dict.Add("AMD Phenom 8400 Triple-Core", 1652);
            dict.Add("AMD Phenom 8450 Triple-Core", 1886);
            dict.Add("AMD Phenom 8450e Triple-Core", 2035);
            dict.Add("AMD Phenom 8600 Triple-Core", 1697);
            dict.Add("AMD Phenom 8600B Triple-Core", 1939);
            dict.Add("AMD Phenom 8650 Triple-Core", 2002);
            dict.Add("AMD Phenom 8750 Triple-Core", 2071);
            dict.Add("AMD Phenom 8750B Triple-Core", 2077);
            dict.Add("AMD Phenom 8850 Triple-Core", 2109);
            dict.Add("AMD Phenom 8850B Triple-Core", 2080);
            dict.Add("AMD Phenom 9100e Quad-Core", 1851);
            dict.Add("AMD Phenom 9150e Quad-Core", 2125);
            dict.Add("AMD Phenom 9350e Quad-Core", 2320);
            dict.Add("AMD Phenom 9450e Quad-Core", 2621);
            dict.Add("AMD Phenom 9500 Quad-Core", 2216);
            dict.Add("AMD Phenom 9550 Quad-Core", 2555);
            dict.Add("AMD Phenom 9600 Quad-Core", 2306);
            dict.Add("AMD Phenom 9600B Quad-Core", 2577);
            dict.Add("AMD Phenom 9650 Quad-Core", 2636);
            dict.Add("AMD Phenom 9750 Quad-Core", 2827);
            dict.Add("AMD Phenom 9750B Quad-Core", 2709);
            dict.Add("AMD Phenom 9850 Quad-Core", 2894);
            dict.Add("AMD Phenom 9850B Quad-Core", 3160);
            dict.Add("AMD Phenom 9950 Quad-Core", 3034);
            dict.Add("AMD Phenom FX-5000 Quad-Core", 2518);
            dict.Add("AMD Phenom FX-5200 Quad-Core", 2948);
            dict.Add("AMD Phenom FX-7750 Quad-Core", 3616);
            dict.Add("AMD Phenom II 42 TWKR Black Edition", 3489);
            dict.Add("AMD Phenom II N620 Dual-Core", 1675);
            dict.Add("AMD Phenom II N640 Dual-Core", 1733);
            dict.Add("AMD Phenom II N660 Dual-Core", 1853);
            dict.Add("AMD Phenom II N830 3 + 1", 2652);
            dict.Add("AMD Phenom II N830 Triple-Core", 1885);
            dict.Add("AMD Phenom II N850 Triple-Core", 1958);
            dict.Add("AMD Phenom II N870 Triple-Core", 1945);
            dict.Add("AMD Phenom II N930 Quad-Core", 2273);
            dict.Add("AMD Phenom II N950 Quad-Core", 2498);
            dict.Add("AMD Phenom II N970 Quad-Core", 2523);
            dict.Add("AMD Phenom II P650 Dual-Core", 1603);
            dict.Add("AMD Phenom II P820 Triple-Core", 1598);
            dict.Add("AMD Phenom II P840 Triple-Core", 1731);
            dict.Add("AMD Phenom II P860 Triple-Core", 1768);
            dict.Add("AMD Phenom II P920 Quad-Core", 1880);
            dict.Add("AMD Phenom II P940 Quad-Core", 1808);
            dict.Add("AMD Phenom II P960 Quad-Core", 2100);
            dict.Add("AMD Phenom II X2 511", 2104);
            dict.Add("AMD Phenom II X2 521", 2090);
            dict.Add("AMD Phenom II X2 545", 1890);
            dict.Add("AMD Phenom II X2 550", 2032);
            dict.Add("AMD Phenom II X2 555", 2045);
            dict.Add("AMD Phenom II X2 560", 2050);
            dict.Add("AMD Phenom II X2 565", 2189);
            dict.Add("AMD Phenom II X2 570", 1960);
            dict.Add("AMD Phenom II X2 B53", 1646);
            dict.Add("AMD Phenom II X2 B55", 1898);
            dict.Add("AMD Phenom II X2 B57", 1884);
            dict.Add("AMD Phenom II X2 B59", 2171);
            dict.Add("AMD Phenom II X3 700e", 2249);
            dict.Add("AMD Phenom II X3 705e", 2244);
            dict.Add("AMD Phenom II X3 710", 2452);
            dict.Add("AMD Phenom II X3 715", 2649);
            dict.Add("AMD Phenom II X3 720", 2691);
            dict.Add("AMD Phenom II X3 740", 3002);
            dict.Add("AMD Phenom II X3 B73", 2749);
            dict.Add("AMD Phenom II X3 B75", 3155);
            dict.Add("AMD Phenom II X3 B77", 3288);
            dict.Add("AMD Phenom II X4 805", 2999);
            dict.Add("AMD Phenom II X4 810", 3139);
            dict.Add("AMD Phenom II X4 820", 3434);
            dict.Add("AMD Phenom II X4 830", 3469);
            dict.Add("AMD Phenom II X4 840", 3520);
            dict.Add("AMD Phenom II X4 840T", 3617);
            dict.Add("AMD Phenom II X4 850", 3539);
            dict.Add("AMD Phenom II X4 900e", 2853);
            dict.Add("AMD Phenom II X4 905e", 3080);
            dict.Add("AMD Phenom II X4 910", 3279);
            dict.Add("AMD Phenom II X4 910e", 3344);
            dict.Add("AMD Phenom II X4 920", 3420);
            dict.Add("AMD Phenom II X4 925", 3439);
            dict.Add("AMD Phenom II X4 940", 3615);
            dict.Add("AMD Phenom II X4 945", 3693);
            dict.Add("AMD Phenom II X4 955", 3969);
            dict.Add("AMD Phenom II X4 960T", 3811);
            dict.Add("AMD Phenom II X4 965", 4221);
            dict.Add("AMD Phenom II X4 970", 4384);
            dict.Add("AMD Phenom II X4 973", 3759);
            dict.Add("AMD Phenom II X4 975", 4534);
            dict.Add("AMD Phenom II X4 977", 3750);
            dict.Add("AMD Phenom II X4 980", 4528);
            dict.Add("AMD Phenom II X4 8700e", 2377);
            dict.Add("AMD Phenom II X4 B05e", 2781);
            dict.Add("AMD Phenom II X4 B15e", 2736);
            dict.Add("AMD Phenom II X4 B25", 3477);
            dict.Add("AMD Phenom II X4 B35", 3443);
            dict.Add("AMD Phenom II X4 B40", 3422);
            dict.Add("AMD Phenom II X4 B45", 3488);
            dict.Add("AMD Phenom II X4 B50", 3720);
            dict.Add("AMD Phenom II X4 B55", 3885);
            dict.Add("AMD Phenom II X4 B60", 4063);
            dict.Add("AMD Phenom II X4 B65", 4177);
            dict.Add("AMD Phenom II X4 B70", 4261);
            dict.Add("AMD Phenom II X4 B93", 3377);
            dict.Add("AMD Phenom II X4 B95", 3762);
            dict.Add("AMD Phenom II X4 B97", 3941);
            dict.Add("AMD Phenom II X4 B99", 4206);
            dict.Add("AMD Phenom II X6 1035T", 4715);
            dict.Add("AMD Phenom II X6 1045T", 4856);
            dict.Add("AMD Phenom II X6 1055T", 5011);
            dict.Add("AMD Phenom II X6 1065T", 5143);
            dict.Add("AMD Phenom II X6 1075T", 5386);
            dict.Add("AMD Phenom II X6 1090T", 5629);
            dict.Add("AMD Phenom II X6 1100T", 5783);
            dict.Add("AMD Phenom II X620 Dual-Core", 2189);
            dict.Add("AMD Phenom II X640 Dual-Core", 2021);
            dict.Add("AMD Phenom II X920 Quad-Core", 2638);
            dict.Add("AMD Phenom II X940 Quad-Core", 2695);
            dict.Add("AMD Phenom Ultra X4 24500", 3945);
            dict.Add("AMD Phenom X2 Dual-Core GE-4060", 1697);
            dict.Add("AMD Phenom X2 Dual-Core GE-5560", 2206);
            dict.Add("AMD Phenom X2 Dual-Core GE-7060", 1752);
            dict.Add("AMD Phenom X2 Dual-Core GP-5000", 1180);
            dict.Add("AMD Phenom X2 Dual-Core GP-7730", 1435);
            dict.Add("AMD Phenom X3 8550", 1930);
            dict.Add("AMD Phenom X4 Quad-Core GP-9500", 1757);
            dict.Add("AMD Phenom X4 Quad-Core GP-9530", 1994);
            dict.Add("AMD Phenom X4 Quad-Core GP-9600", 2235);
            dict.Add("AMD Phenom X4 Quad-Core GP-9730", 3033);
            dict.Add("AMD Phenom X4 Quad-Core GP-9830", 2715);
            dict.Add("AMD Phenom X4 Quad-Core GP-9930", 2690);
            dict.Add("AMD Phenom X4 Quad-Core GS-5560", 1211);
            dict.Add("AMD PRO A4-3350B APU", 2691);
            dict.Add("AMD PRO A4-8350B", 2881);
            dict.Add("AMD PRO A6-8500B", 2516);
            dict.Add("AMD PRO A6-8530B", 2452);
            dict.Add("AMD PRO A6-8550B", 3042);
            dict.Add("AMD PRO A6-8570", 3264);
            dict.Add("AMD PRO A6-8570E", 2906);
            dict.Add("AMD PRO A6-9500", 3262);
            dict.Add("AMD PRO A6-9500B", 2382);
            dict.Add("AMD PRO A6-9500E", 2752);
            dict.Add("AMD PRO A8-8600B", 3880);
            dict.Add("AMD PRO A8-8650B", 5476);
            dict.Add("AMD PRO A8-9600", 5336);
            dict.Add("AMD PRO A8-9600B", 3970);
            dict.Add("AMD PRO A10-8700B", 3725);
            dict.Add("AMD PRO A10-8730B", 4075);
            dict.Add("AMD PRO A10-8750B", 5653);
            dict.Add("AMD PRO A10-8770", 5742);
            dict.Add("AMD PRO A10-8770E", 4901);
            dict.Add("AMD PRO A10-8850B", 5848);
            dict.Add("AMD PRO A10-9700", 5742);
            dict.Add("AMD PRO A10-9700E", 5221);
            dict.Add("AMD PRO A12-8800B", 4247);
            dict.Add("AMD PRO A12-8870", 6259);
            dict.Add("AMD PRO A12-8870E", 4735);
            dict.Add("AMD PRO A12-9800", 6205);
            dict.Add("AMD PRO A12-9800B", 4284);
            dict.Add("AMD PRO A12-9800E", 4835);
            dict.Add("AMD QC-4000", 1605);
            dict.Add("AMD R-260H APU", 1576);
            dict.Add("AMD R-460L APU", 1954);
            dict.Add("AMD R-464L APU", 2976);
            dict.Add("AMD RX-427BB", 4468);
            dict.Add("AMD Ryzen 5 1400", 8691);
            dict.Add("AMD Ryzen 5 1500X", 10850);
            dict.Add("AMD Ryzen 5 1600", 12526);
            dict.Add("AMD Ryzen 5 1600X", 13150);
            dict.Add("AMD Ryzen 7 1700", 13804);
            dict.Add("AMD Ryzen 7 1700X", 14717);
            dict.Add("AMD Ryzen 7 1800X", 15368);
            dict.Add("AMD Sempron 130", 656);
            dict.Add("AMD Sempron 140", 736);
            dict.Add("AMD Sempron 145", 802);
            dict.Add("AMD Sempron 150", 654);
            dict.Add("AMD Sempron 200U", 267);
            dict.Add("AMD Sempron 210U", 382);
            dict.Add("AMD Sempron 2100 +", 392);
            dict.Add("AMD Sempron 2200 +", 281);
            dict.Add("AMD Sempron 2300 +", 318);
            dict.Add("AMD Sempron 2400 +", 313);
            dict.Add("AMD Sempron 2500 +", 345);
            dict.Add("AMD Sempron 2600 +", 373);
            dict.Add("AMD Sempron 2650 APU", 903);
            dict.Add("AMD Sempron 2800 +", 384);
            dict.Add("AMD Sempron 3000 +", 408);
            dict.Add("AMD Sempron 3100 +", 452);
            dict.Add("AMD Sempron 3200 +", 401);
            dict.Add("AMD Sempron 3300 +", 484);
            dict.Add("AMD Sempron 3400 +", 449);
            dict.Add("AMD Sempron 3500 +", 402);
            dict.Add("AMD Sempron 3600 +", 489);
            dict.Add("AMD Sempron 3800 +", 554);
            dict.Add("AMD Sempron 3850 APU", 1686);
            dict.Add("AMD Sempron Dual Core 2100", 869);
            dict.Add("AMD Sempron Dual Core 2200", 995);
            dict.Add("AMD Sempron Dual Core 2300", 982);
            dict.Add("AMD Sempron Dual Core 4700", 948);
            dict.Add("AMD Sempron Dual Core 4900", 1081);
            dict.Add("AMD Sempron LE-250", 490);
            dict.Add("AMD Sempron LE-1100", 428);
            dict.Add("AMD Sempron LE-1150", 434);
            dict.Add("AMD Sempron LE-1200", 524);
            dict.Add("AMD Sempron LE-1250", 540);
            dict.Add("AMD Sempron LE-1300", 620);
            dict.Add("AMD Sempron LE-1600", 535);
            dict.Add("AMD Sempron LE-1620", 635);
            dict.Add("AMD Sempron LE-1640", 759);
            dict.Add("AMD Sempron M100", 554);
            dict.Add("AMD Sempron M120", 571);
            dict.Add("AMD Sempron SI-40", 449);
            dict.Add("AMD Sempron SI-42", 451);
            dict.Add("AMD Sempron X2 180", 1222);
            dict.Add("AMD Sempron X2 190", 1529);
            dict.Add("AMD Sempron X2 198 Dual-Core", 1413);
            dict.Add("AMD Turion 64 Mobile MK-36", 462);
            dict.Add("AMD Turion 64 Mobile MK-38", 505);
            dict.Add("AMD Turion 64 Mobile ML-28", 390);
            dict.Add("AMD Turion 64 Mobile ML-30", 403);
            dict.Add("AMD Turion 64 Mobile ML-32", 410);
            dict.Add("AMD Turion 64 Mobile ML-34", 425);
            dict.Add("AMD Turion 64 Mobile ML-37", 447);
            dict.Add("AMD Turion 64 Mobile ML-40", 520);
            dict.Add("AMD Turion 64 Mobile ML-42", 615);
            dict.Add("AMD Turion 64 Mobile ML-44", 591);
            dict.Add("AMD Turion 64 Mobile MT-28", 419);
            dict.Add("AMD Turion 64 Mobile MT-30", 378);
            dict.Add("AMD Turion 64 Mobile MT-32", 426);
            dict.Add("AMD Turion 64 Mobile MT-34", 429);
            dict.Add("AMD Turion 64 Mobile MT-37", 484);
            dict.Add("AMD Turion 64 Mobile MT-40", 536);
            dict.Add("AMD Turion 64 X2 Mobile TL-50", 752);
            dict.Add("AMD Turion 64 X2 Mobile TL-52", 756);
            dict.Add("AMD Turion 64 X2 Mobile TL-56", 876);
            dict.Add("AMD Turion 64 X2 Mobile TL-58", 950);
            dict.Add("AMD Turion 64 X2 Mobile TL-60", 987);
            dict.Add("AMD Turion 64 X2 Mobile TL-62", 1093);
            dict.Add("AMD Turion 64 X2 Mobile TL-64", 1120);
            dict.Add("AMD Turion 64 X2 Mobile TL-66", 1122);
            dict.Add("AMD Turion 64 X2 Mobile TL-68", 1149);
            dict.Add("AMD Turion Dual-Core RM-70", 895);
            dict.Add("AMD Turion Dual-Core RM-72", 839);
            dict.Add("AMD Turion Dual-Core RM-74", 1016);
            dict.Add("AMD Turion Dual-Core RM-75", 925);
            dict.Add("AMD Turion II Dual-Core Mobile M500", 1306);
            dict.Add("AMD Turion II Dual-Core Mobile M520", 1303);
            dict.Add("AMD Turion II Dual-Core Mobile M540", 1295);
            dict.Add("AMD Turion II N530 Dual-Core", 1583);
            dict.Add("AMD Turion II N550 Dual-Core", 1521);
            dict.Add("AMD Turion II Neo K625 Dual-Core", 973);
            dict.Add("AMD Turion II Neo K685 Dual-Core", 1186);
            dict.Add("AMD Turion II Neo N40L Dual-Core", 917);
            dict.Add("AMD Turion II Neo N54L Dual-Core", 1397);
            dict.Add("AMD Turion II P520 Dual-Core", 1363);
            dict.Add("AMD Turion II P540 Dual-Core", 1449);
            dict.Add("AMD Turion II P560 Dual-Core", 1500);
            dict.Add("AMD Turion II Ultra Dual-Core Mobile M600", 1392);
            dict.Add("AMD Turion II Ultra Dual-Core Mobile M620", 1419);
            dict.Add("AMD Turion II Ultra Dual-Core Mobile M640", 1669);
            dict.Add("AMD Turion II Ultra Dual-Core Mobile M660", 1533);
            dict.Add("AMD Turion Neo X2 Dual Core L625", 821);
            dict.Add("AMD Turion X2 Dual Core L510", 871);
            dict.Add("AMD Turion X2 Dual Core Mobile RM-70", 939);
            dict.Add("AMD Turion X2 Dual Core Mobile RM-76", 1041);
            dict.Add("AMD Turion X2 Dual-Core Mobile RM-70", 996);
            dict.Add("AMD Turion X2 Dual-Core Mobile RM-72", 1087);
            dict.Add("AMD Turion X2 Dual-Core Mobile RM-74", 1061);
            dict.Add("AMD Turion X2 Dual-Core Mobile RM-75", 1146);
            dict.Add("AMD Turion X2 Dual-Core Mobile RM-77", 1230);
            dict.Add("AMD Turion X2 Ultra Dual-Core Mobile ZM-80", 1004);
            dict.Add("AMD Turion X2 Ultra Dual-Core Mobile ZM-82", 1081);
            dict.Add("AMD Turion X2 Ultra Dual-Core Mobile ZM-84", 1208);
            dict.Add("AMD Turion X2 Ultra Dual-Core Mobile ZM-85", 1232);
            dict.Add("AMD Turion X2 Ultra Dual-Core Mobile ZM-86", 1279);
            dict.Add("AMD Turion X2 Ultra Dual-Core Mobile ZM-87", 1344);
            dict.Add("AMD TurionX2 Dual Core Mobile RM-70", 979);
            dict.Add("AMD TurionX2 Dual Core Mobile RM-72", 1112);
            dict.Add("AMD TurionX2 Ultra DualCore Mobile ZM-85", 1454);
            dict.Add("AMD V105", 274);
            dict.Add("AMD V120", 631);
            dict.Add("AMD V140", 640);
            dict.Add("AMD V160", 680);
            dict.Add("AMD Z-01", 523);
            dict.Add("AMD Z-60 APU", 485);
            dict.Add("AMD-K6 3D", 90);
            dict.Add("AMD-K6-III", 108);
            dict.Add("AMD-K7", 403);
            dict.Add("Athlon 64 Dual Core 3800 +", 1234);
            dict.Add("Athlon 64 Dual Core 4200 +", 1156);
            dict.Add("Athlon 64 Dual Core 4800 +", 1370);
            dict.Add("Athlon 64 Dual Core 5000 +", 1497);
            dict.Add("Athlon 64 Dual Core 5400 +", 1614);
            dict.Add("Athlon 64 Dual Core 5600 +", 1548);
            dict.Add("Athlon Dual Core 4050e", 968);
            dict.Add("Athlon Dual Core 4450e", 1081);
            dict.Add("Athlon Dual Core 4850e", 1327);
            dict.Add("Celeron Dual-Core Q8300 @ 2.50GHz", 2837);
            dict.Add("Celeron Dual-Core T3000 @ 1.80GHz", 1089);
            dict.Add("Celeron Dual-Core T3100 @ 1.90GHz", 1176);
            dict.Add("Celeron Dual-Core T3300 @ 2.00GHz", 1222);
            dict.Add("Celeron Dual-Core T3500 @ 2.10GHz", 1292);
            dict.Add("Dual-Core AMD Opteron 1220 SE", 1206);
            dict.Add("HP Hexa-Core 2.0GHz", 2178);
            dict.Add("Intel 1300 @ 1.66GHz", 857);
            dict.Add("Intel 1400 @ 1.83GHz", 980);
            dict.Add("Intel 1500 @ 2.00GHz", 1024);
            dict.Add("Intel Atom 230 @ 1.60GHz", 308);
            dict.Add("Intel Atom 330 @ 1.60GHz", 591);
            dict.Add("Intel Atom C2338 @ 1.74GHz", 898);
            dict.Add("Intel Atom C2358 @ 1.74GHz", 961);
            dict.Add("Intel Atom C2550 @ 2.40GHz", 2329);
            dict.Add("Intel Atom C2558 @ 2.40GHz", 2169);
            dict.Add("Intel Atom C2750 @ 2.40GHz", 3800);
            dict.Add("Intel Atom C2750 @ 2.41GHz", 3344);
            dict.Add("Intel Atom C2758 @ 2.40GHz", 3162);
            dict.Add("Intel Atom D410 @ 1.66GHz", 310);
            dict.Add("Intel Atom D425 @ 1.80GHz", 343);
            dict.Add("Intel Atom D510 @ 1.66GHz", 654);
            dict.Add("Intel Atom D525 @ 1.80GHz", 701);
            dict.Add("Intel Atom D2500 @ 1.86GHz", 406);
            dict.Add("Intel Atom D2550 @ 1.86GHz", 672);
            dict.Add("Intel Atom D2560 @ 2.00GHz", 711);
            dict.Add("Intel Atom D2700 @ 2.13GHz", 840);
            dict.Add("Intel Atom D2701 @ 2.13GHz", 720);
            dict.Add("Intel Atom E660 @ 1.30GHz", 271);
            dict.Add("Intel Atom E680 @ 1.60GHz", 346);
            dict.Add("Intel Atom E3815 @ 1.46GHz", 362);
            dict.Add("Intel Atom E3825 @ 1.33GHz", 564);
            dict.Add("Intel Atom E3826 @ 1.46GHz", 402);
            dict.Add("Intel Atom E3827 @ 1.74GHz", 849);
            dict.Add("Intel Atom E3840 @ 1.91GHz", 1162);
            dict.Add("Intel Atom E3845 @ 1.91GHz", 1466);
            dict.Add("Intel Atom E3950 @ 1.60GHz", 2034);
            dict.Add("Intel Atom K510 @ 1.66GHz", 658);
            dict.Add("Intel Atom N270 @ 1.60GHz", 270);
            dict.Add("Intel Atom N280 @ 1.66GHz", 289);
            dict.Add("Intel Atom N435 @ 1.33GHz", 236);
            dict.Add("Intel Atom N450 @ 1.66GHz", 297);
            dict.Add("Intel Atom N455 @ 1.66GHz", 288);
            dict.Add("Intel Atom N470 @ 1.83GHz", 307);
            dict.Add("Intel Atom N475 @ 1.83GHz", 314);
            dict.Add("Intel Atom N550 @ 1.50GHz", 516);
            dict.Add("Intel Atom N570 @ 1.66GHz", 579);
            dict.Add("Intel Atom N2100 @ 1.60GHz", 309);
            dict.Add("Intel Atom N2600 @ 1.60GHz", 528);
            dict.Add("Intel Atom N2800 @ 1.86GHz", 619);
            dict.Add("Intel Atom S1260 @ 2.00GHz", 916);
            dict.Add("Intel Atom x5-E3930 @ 1.30GHz", 1021);
            dict.Add("Intel Atom x5-E8000 @ 1.04GHz", 1605);
            dict.Add("Intel Atom x5-Z8300 @ 1.44GHz", 1200);
            dict.Add("Intel Atom x5-Z8330 @ 1.44GHz", 1006);
            dict.Add("Intel Atom x5-Z8350 @ 1.44GHz", 1261);
            dict.Add("Intel Atom x5-Z8500 @ 1.44GHz", 1695);
            dict.Add("Intel Atom x5-Z8550 @ 1.44GHz", 1855);
            dict.Add("Intel Atom x7-Z8700 @ 1.60GHz", 1923);
            dict.Add("Intel Atom x7-Z8750 @ 1.60GHz", 1853);
            dict.Add("Intel Atom Z510 @ 1.10GHz", 194);
            dict.Add("Intel Atom Z515 @ 1.20GHz", 218);
            dict.Add("Intel Atom Z520 @ 1.33GHz", 215);
            dict.Add("Intel Atom Z530 @ 1.60GHz", 281);
            dict.Add("Intel Atom Z540 @ 1.86GHz", 348);
            dict.Add("Intel Atom Z550 @ 2.00GHz", 381);
            dict.Add("Intel Atom Z670 @ 1.50GHz", 249);
            dict.Add("Intel Atom Z2760 @ 1.80GHz", 571);
            dict.Add("Intel Atom Z3735D @ 1.33GHz", 912);
            dict.Add("Intel Atom Z3735E @ 1.33GHz", 927);
            dict.Add("Intel Atom Z3735F @ 1.33GHz", 911);
            dict.Add("Intel Atom Z3735G @ 1.33GHz", 909);
            dict.Add("Intel Atom Z3736F @ 1.33GHz", 917);
            dict.Add("Intel Atom Z3740 @ 1.33GHz", 1061);
            dict.Add("Intel Atom Z3740D @ 1.33GHz", 1020);
            dict.Add("Intel Atom Z3745 @ 1.33GHz", 1071);
            dict.Add("Intel Atom Z3745D @ 1.33GHz", 1082);
            dict.Add("Intel Atom Z3770 @ 1.46GHz", 1239);
            dict.Add("Intel Atom Z3770D @ 1.49GHz", 805);
            dict.Add("Intel Atom Z3775 @ 1.46GHz", 1253);
            dict.Add("Intel Atom Z3775D @ 1.49GHz", 1278);
            dict.Add("Intel Atom Z3795 @ 1.60GHz", 1767);
            dict.Add("Intel Celeron 1.70GHz", 130);
            dict.Add("Intel Celeron 1.80GHz", 136);
            dict.Add("Intel Celeron 2.00GHz", 184);
            dict.Add("Intel Celeron 2.10GHz", 244);
            dict.Add("Intel Celeron 2.13GHz", 210);
            dict.Add("Intel Celeron 2.20GHz", 213);
            dict.Add("Intel Celeron 2.26GHz", 223);
            dict.Add("Intel Celeron 2.30GHz", 227);
            dict.Add("Intel Celeron 2.40GHz", 218);
            dict.Add("Intel Celeron 2.50GHz", 253);
            dict.Add("Intel Celeron 2.53GHz", 246);
            dict.Add("Intel Celeron 2.60GHz", 241);
            dict.Add("Intel Celeron 2.66GHz", 251);
            dict.Add("Intel Celeron 2.70GHz", 245);
            dict.Add("Intel Celeron 2.80GHz", 255);
            dict.Add("Intel Celeron 2.93GHz", 285);
            dict.Add("Intel Celeron 3.06GHz", 286);
            dict.Add("Intel Celeron 3.20GHz", 289);
            dict.Add("Intel Celeron 3.33GHz", 333);
            dict.Add("Intel Celeron 215 @ 1.33GHz", 322);
            dict.Add("Intel Celeron 220 @ 1.20GHz", 320);
            dict.Add("Intel Celeron 420 @ 1.60GHz", 455);
            dict.Add("Intel Celeron 430 @ 1.80GHz", 491);
            dict.Add("Intel Celeron 440 @ 2.00GHz", 537);
            dict.Add("Intel Celeron 450 @ 2.20GHz", 621);
            dict.Add("Intel Celeron 530 @ 1.73GHz", 478);
            dict.Add("Intel Celeron 540 @ 1.86GHz", 509);
            dict.Add("Intel Celeron 550 @ 2.00GHz", 501);
            dict.Add("Intel Celeron 560 @ 2.13GHz", 532);
            dict.Add("Intel Celeron 570 @ 2.26GHz", 526);
            dict.Add("Intel Celeron 600MHz", 147);
            dict.Add("Intel Celeron 723 @ 1.20GHz", 404);
            dict.Add("Intel Celeron 743 @ 1.30GHz", 418);
            dict.Add("Intel Celeron 807 @ 1.50GHz", 751);
            dict.Add("Intel Celeron 807UE @ 1.00GHz", 369);
            dict.Add("Intel Celeron 827E @ 1.40GHz", 613);
            dict.Add("Intel Celeron 847 @ 1.10GHz", 937);
            dict.Add("Intel Celeron 847E @ 1.10GHz", 1026);
            dict.Add("Intel Celeron 857 @ 1.20GHz", 1170);
            dict.Add("Intel Celeron 867 @ 1.30GHz", 1199);
            dict.Add("Intel Celeron 877 @ 1.40GHz", 1301);
            dict.Add("Intel Celeron 887 @ 1.50GHz", 1297);
            dict.Add("Intel Celeron 900 @ 2.20GHz", 678);
            dict.Add("Intel Celeron 925 @ 2.30GHz", 691);
            dict.Add("Intel Celeron 1000M @ 1.80GHz", 1659);
            dict.Add("Intel Celeron 1000MHz", 241);
            dict.Add("Intel Celeron 1005M @ 1.90GHz", 1850);
            dict.Add("Intel Celeron 1007U @ 1.50GHz", 1399);
            dict.Add("Intel Celeron 1017U @ 1.60GHz", 1563);
            dict.Add("Intel Celeron 1019Y @ 1.00GHz", 960);
            dict.Add("Intel Celeron 1020E @ 2.20GHz", 2232);
            dict.Add("Intel Celeron 1037U @ 1.80GHz", 1740);
            dict.Add("Intel Celeron 1047UE @ 1.40GHz", 1188);
            dict.Add("Intel Celeron 1066MHz", 259);
            dict.Add("Intel Celeron 1100MHz", 174);
            dict.Add("Intel Celeron 1133MHz", 282);
            dict.Add("Intel Celeron 1200MHz", 244);
            dict.Add("Intel Celeron 1300MHz", 288);
            dict.Add("Intel Celeron 1333MHz", 290);
            dict.Add("Intel Celeron 1400MHz", 302);
            dict.Add("Intel Celeron 2000E @ 2.20GHz", 2365);
            dict.Add("Intel Celeron 2950M @ 2.00GHz", 2043);
            dict.Add("Intel Celeron 2955U @ 1.40GHz", 1449);
            dict.Add("Intel Celeron 2957U @ 1.40GHz", 1464);
            dict.Add("Intel Celeron 2961Y @ 1.10GHz", 1110);
            dict.Add("Intel Celeron 2970M @ 2.20GHz", 2305);
            dict.Add("Intel Celeron 2980U @ 1.60GHz", 1563);
            dict.Add("Intel Celeron 2981U @ 1.60GHz", 1542);
            dict.Add("Intel Celeron 3205U @ 1.50GHz", 1665);
            dict.Add("Intel Celeron 3215U @ 1.70GHz", 1766);
            dict.Add("Intel Celeron 3755U @ 1.70GHz", 1702);
            dict.Add("Intel Celeron 3765U @ 1.90GHz", 2018);
            dict.Add("Intel Celeron 3855U @ 1.60GHz", 1694);
            dict.Add("Intel Celeron 3955U @ 2.00GHz", 1707);
            dict.Add("Intel Celeron @ 1.30GHz", 965);
            dict.Add("Intel Celeron B710 @ 1.60GHz", 186);
            dict.Add("Intel Celeron B720 @ 1.70GHz", 814);
            dict.Add("Intel Celeron B800 @ 1.50GHz", 1287);
            dict.Add("Intel Celeron B810 @ 1.60GHz", 1373);
            dict.Add("Intel Celeron B815 @ 1.60GHz", 1363);
            dict.Add("Intel Celeron B820 @ 1.70GHz", 1435);
            dict.Add("Intel Celeron B830 @ 1.80GHz", 1528);
            dict.Add("Intel Celeron B840 @ 1.90GHz", 1739);
            dict.Add("Intel Celeron D 347 @ 3.06GHz", 277);
            dict.Add("Intel Celeron D 352 @ 3.20GHz", 305);
            dict.Add("Intel Celeron D 356 @ 3.33GHz", 315);
            dict.Add("Intel Celeron D 360 @ 3.46GHz", 334);
            dict.Add("Intel Celeron D 365 @ 3.60GHz", 404);
            dict.Add("Intel Celeron D 420 @ 1.60GHz", 481);
            dict.Add("Intel Celeron D 430 @ 1.80GHz", 585);
            dict.Add("Intel Celeron E1200 @ 1.60GHz", 861);
            dict.Add("Intel Celeron E1400 @ 2.00GHz", 1042);
            dict.Add("Intel Celeron E1500 @ 2.20GHz", 1143);
            dict.Add("Intel Celeron E1600 @ 2.40GHz", 1260);
            dict.Add("Intel Celeron E3200 @ 2.40GHz", 1379);
            dict.Add("Intel Celeron E3300 @ 2.50GHz", 1386);
            dict.Add("Intel Celeron E3400 @ 2.60GHz", 1425);
            dict.Add("Intel Celeron E3500 @ 2.70GHz", 1422);
            dict.Add("Intel Celeron G440 @ 1.60GHz", 698);
            dict.Add("Intel Celeron G460 @ 1.80GHz", 1023);
            dict.Add("Intel Celeron G465 @ 1.90GHz", 1030);
            dict.Add("Intel Celeron G470 @ 2.00GHz", 1305);
            dict.Add("Intel Celeron G530 @ 2.40GHz", 2147);
            dict.Add("Intel Celeron G530T @ 2.00GHz", 1604);
            dict.Add("Intel Celeron G540 @ 2.50GHz", 2203);
            dict.Add("Intel Celeron G540T @ 2.10GHz", 2114);
            dict.Add("Intel Celeron G550 @ 2.60GHz", 2299);
            dict.Add("Intel Celeron G550T @ 2.20GHz", 2073);
            dict.Add("Intel Celeron G555 @ 2.70GHz", 2422);
            dict.Add("Intel Celeron G1101 @ 2.27GHz", 1727);
            dict.Add("Intel Celeron G1610 @ 2.60GHz", 2503);
            dict.Add("Intel Celeron G1610T @ 2.30GHz", 2322);
            dict.Add("Intel Celeron G1620 @ 2.70GHz", 2590);
            dict.Add("Intel Celeron G1620T @ 2.40GHz", 2287);
            dict.Add("Intel Celeron G1630 @ 2.80GHz", 2520);
            dict.Add("Intel Celeron G1820 @ 2.70GHz", 2778);
            dict.Add("Intel Celeron G1820T @ 2.40GHz", 2495);
            dict.Add("Intel Celeron G1820TE @ 2.20GHz", 2230);
            dict.Add("Intel Celeron G1830 @ 2.80GHz", 2614);
            dict.Add("Intel Celeron G1840 @ 2.80GHz", 2996);
            dict.Add("Intel Celeron G1840T @ 2.50GHz", 2613);
            dict.Add("Intel Celeron G1850 @ 2.90GHz", 2933);
            dict.Add("Intel Celeron G3900 @ 2.80GHz", 3395);
            dict.Add("Intel Celeron G3900E @ 2.40GHz", 2707);
            dict.Add("Intel Celeron G3900T @ 2.60GHz", 2944);
            dict.Add("Intel Celeron G3920 @ 2.90GHz", 3477);
            dict.Add("Intel Celeron G3930 @ 2.90GHz", 2960);
            dict.Add("Intel Celeron G3950 @ 3.00GHz", 3334);
            dict.Add("Intel Celeron J1750 @ 2.41GHz", 1037);
            dict.Add("Intel Celeron J1800 @ 2.41GHz", 1027);
            dict.Add("Intel Celeron J1850 @ 1.99GHz", 1614);
            dict.Add("Intel Celeron J1900 @ 1.99GHz", 1863);
            dict.Add("Intel Celeron J3060 @ 1.60GHz", 993);
            dict.Add("Intel Celeron J3160 @ 1.60GHz", 1839);
            dict.Add("Intel Celeron J3355 @ 2.00GHz", 1333);
            dict.Add("Intel Celeron J3455 @ 1.50GHz", 2153);
            dict.Add("Intel Celeron M 1.00GHz", 207);
            dict.Add("Intel Celeron M 1.30GHz", 316);
            dict.Add("Intel Celeron M 1.50GHz", 346);
            dict.Add("Intel Celeron M 1.60GHz", 365);
            dict.Add("Intel Celeron M 1.70GHz", 393);
            dict.Add("Intel Celeron M 360 1.40GHz", 327);
            dict.Add("Intel Celeron M 410 @ 1.46GHz", 317);
            dict.Add("Intel Celeron M 420 @ 1.60GHz", 339);
            dict.Add("Intel Celeron M 430 @ 1.73GHz", 369);
            dict.Add("Intel Celeron M 440 @ 1.86GHz", 414);
            dict.Add("Intel Celeron M 443 @ 1.20GHz", 253);
            dict.Add("Intel Celeron M 450 @ 2.00GHz", 455);
            dict.Add("Intel Celeron M 520 @ 1.60GHz", 432);
            dict.Add("Intel Celeron M 530 @ 1.73GHz", 461);
            dict.Add("Intel Celeron M 540 @ 1.86GHz", 542);
            dict.Add("Intel Celeron M 600MHz", 151);
            dict.Add("Intel Celeron M 723 @ 1.20GHz", 457);
            dict.Add("Intel Celeron M 900MHz", 207);
            dict.Add("Intel Celeron M 1200MHz", 275);
            dict.Add("Intel Celeron M 1300MHz", 303);
            dict.Add("Intel Celeron M 1500MHz", 332);
            dict.Add("Intel Celeron M ULV 800MHz", 184);
            dict.Add("Intel Celeron N2805 @ 1.46GHz", 453);
            dict.Add("Intel Celeron N2806 @ 1.60GHz", 773);
            dict.Add("Intel Celeron N2807 @ 1.58GHz", 843);
            dict.Add("Intel Celeron N2808 @ 1.58GHz", 938);
            dict.Add("Intel Celeron N2810 @ 2.00GHz", 791);
            dict.Add("Intel Celeron N2815 @ 1.86GHz", 855);
            dict.Add("Intel Celeron N2820 @ 2.13GHz", 970);
            dict.Add("Intel Celeron N2830 @ 2.16GHz", 959);
            dict.Add("Intel Celeron N2840 @ 2.16GHz", 1010);
            dict.Add("Intel Celeron N2910 @ 1.60GHz", 1231);
            dict.Add("Intel Celeron N2920 @ 1.86GHz", 1522);
            dict.Add("Intel Celeron N2930 @ 1.83GHz", 1631);
            dict.Add("Intel Celeron N2940 @ 1.83GHz", 1753);
            dict.Add("Intel Celeron N3000 @ 1.04GHz", 917);
            dict.Add("Intel Celeron N3010 @ 1.04GHz", 962);
            dict.Add("Intel Celeron N3050 @ 1.60GHz", 889);
            dict.Add("Intel Celeron N3060 @ 1.60GHz", 973);
            dict.Add("Intel Celeron N3150 @ 1.60GHz", 1678);
            dict.Add("Intel Celeron N3160 @ 1.60GHz", 1697);
            dict.Add("Intel Celeron N3350 @ 1.10GHz", 1142);
            dict.Add("Intel Celeron N3450 @ 1.10GHz", 1873);
            dict.Add("Intel Celeron P4500 @ 1.87GHz", 1152);
            dict.Add("Intel Celeron P4505 @ 1.87GHz", 994);
            dict.Add("Intel Celeron P4600 @ 2.00GHz", 1345);
            dict.Add("Intel Celeron SU2300 @ 1.20GHz", 782);
            dict.Add("Intel Celeron T1600 @ 1.66GHz", 931);
            dict.Add("Intel Celeron T1700 @ 1.83GHz", 1058);
            dict.Add("Intel Celeron U1900 @ 1.99GHz", 1704);
            dict.Add("Intel Celeron U3400 @ 1.07GHz", 727);
            dict.Add("Intel Celeron U3405 @ 1.07GHz", 684);
            dict.Add("Intel Celeron U3600 @ 1.20GHz", 777);
            dict.Add("Intel Core2 Duo E4300 @ 1.80GHz", 1049);
            dict.Add("Intel Core2 Duo E4400 @ 2.00GHz", 1158);
            dict.Add("Intel Core2 Duo E4500 @ 2.20GHz", 1275);
            dict.Add("Intel Core2 Duo E4600 @ 2.40GHz", 1387);
            dict.Add("Intel Core2 Duo E4700 @ 2.60GHz", 1476);
            dict.Add("Intel Core2 Duo E6300 @ 1.86GHz", 1112);
            dict.Add("Intel Core2 Duo E6320 @ 1.86GHz", 1193);
            dict.Add("Intel Core2 Duo E6400 @ 2.13GHz", 1296);
            dict.Add("Intel Core2 Duo E6420 @ 2.13GHz", 1372);
            dict.Add("Intel Core2 Duo E6540 @ 2.33GHz", 1464);
            dict.Add("Intel Core2 Duo E6550 @ 2.33GHz", 1501);
            dict.Add("Intel Core2 Duo E6600 @ 2.40GHz", 1555);
            dict.Add("Intel Core2 Duo E6700 @ 2.66GHz", 1703);
            dict.Add("Intel Core2 Duo E6750 @ 2.66GHz", 1720);
            dict.Add("Intel Core2 Duo E6850 @ 3.00GHz", 1951);
            dict.Add("Intel Core2 Duo E7200 @ 2.53GHz", 1637);
            dict.Add("Intel Core2 Duo E7300 @ 2.66GHz", 1728);
            dict.Add("Intel Core2 Duo E7400 @ 2.80GHz", 1778);
            dict.Add("Intel Core2 Duo E7500 @ 2.93GHz", 1879);
            dict.Add("Intel Core2 Duo E7600 @ 3.06GHz", 2013);
            dict.Add("Intel Core2 Duo E8135 @ 2.40GHz", 1681);
            dict.Add("Intel Core2 Duo E8135 @ 2.66GHz", 1765);
            dict.Add("Intel Core2 Duo E8200 @ 2.66GHz", 1883);
            dict.Add("Intel Core2 Duo E8235 @ 2.80GHz", 1940);
            dict.Add("Intel Core2 Duo E8290 @ 2.83GHz", 2353);
            dict.Add("Intel Core2 Duo E8300 @ 2.83GHz", 1996);
            dict.Add("Intel Core2 Duo E8335 @ 2.66GHz", 1810);
            dict.Add("Intel Core2 Duo E8335 @ 2.93GHz", 2163);
            dict.Add("Intel Core2 Duo E8400 @ 3.00GHz", 2163);
            dict.Add("Intel Core2 Duo E8435 @ 3.06GHz", 2139);
            dict.Add("Intel Core2 Duo E8500 @ 3.16GHz", 2295);
            dict.Add("Intel Core2 Duo E8600 @ 3.33GHz", 2417);
            dict.Add("Intel Core2 Duo E8700 @ 3.50GHz", 2212);
            dict.Add("Intel Core2 Duo L7100 @ 1.20GHz", 727);
            dict.Add("Intel Core2 Duo L7200 @ 1.33GHz", 684);
            dict.Add("Intel Core2 Duo L7300 @ 1.40GHz", 841);
            dict.Add("Intel Core2 Duo L7400 @ 1.50GHz", 846);
            dict.Add("Intel Core2 Duo L7500 @ 1.60GHz", 965);
            dict.Add("Intel Core2 Duo L7700 @ 1.80GHz", 965);
            dict.Add("Intel Core2 Duo L7800 @ 2.00GHz", 1116);
            dict.Add("Intel Core2 Duo L9300 @ 1.60GHz", 1090);
            dict.Add("Intel Core2 Duo L9600 @ 2.13GHz", 1348);
            dict.Add("Intel Core2 Duo P7350 @ 2.00GHz", 1301);
            dict.Add("Intel Core2 Duo P7370 @ 2.00GHz", 1301);
            dict.Add("Intel Core2 Duo P7450 @ 2.13GHz", 1428);
            dict.Add("Intel Core2 Duo P7500 @ 1.60GHz", 945);
            dict.Add("Intel Core2 Duo P7550 @ 2.26GHz", 1523);
            dict.Add("Intel Core2 Duo P7570 @ 2.26GHz", 1447);
            dict.Add("Intel Core2 Duo P7700 @ 1.80GHz", 1021);
            dict.Add("Intel Core2 Duo P8400 @ 2.26GHz", 1462);
            dict.Add("Intel Core2 Duo P8600 @ 2.40GHz", 1542);
            dict.Add("Intel Core2 Duo P8700 @ 2.53GHz", 1639);
            dict.Add("Intel Core2 Duo P8800 @ 2.66GHz", 1751);
            dict.Add("Intel Core2 Duo P9300 @ 2.26GHz", 1583);
            dict.Add("Intel Core2 Duo P9500 @ 2.53GHz", 1789);
            dict.Add("Intel Core2 Duo P9600 @ 2.53GHz", 1768);
            dict.Add("Intel Core2 Duo P9600 @ 2.66GHz", 1873);
            dict.Add("Intel Core2 Duo P9700 @ 2.80GHz", 2025);
            dict.Add("Intel Core2 Duo Q6867 @ 3.00GHz", 589);
            dict.Add("Intel Core2 Duo SL9400 @ 1.86GHz", 1271);
            dict.Add("Intel Core2 Duo SP9400 @ 2.40GHz", 1679);
            dict.Add("Intel Core2 Duo SU9400 @ 1.40GHz", 936);
            dict.Add("Intel Core2 Duo T5200 @ 1.60GHz", 835);
            dict.Add("Intel Core2 Duo T5250 @ 1.50GHz", 830);
            dict.Add("Intel Core2 Duo T5270 @ 1.40GHz", 821);
            dict.Add("Intel Core2 Duo T5300 @ 1.73GHz", 940);
            dict.Add("Intel Core2 Duo T5450 @ 1.66GHz", 906);
            dict.Add("Intel Core2 Duo T5470 @ 1.60GHz", 906);
            dict.Add("Intel Core2 Duo T5500 @ 1.66GHz", 921);
            dict.Add("Intel Core2 Duo T5550 @ 1.83GHz", 1041);
            dict.Add("Intel Core2 Duo T5600 @ 1.83GHz", 1025);
            dict.Add("Intel Core2 Duo T5670 @ 1.80GHz", 1009);
            dict.Add("Intel Core2 Duo T5750 @ 2.00GHz", 1091);
            dict.Add("Intel Core2 Duo T5800 @ 2.00GHz", 1112);
            dict.Add("Intel Core2 Duo T5850 @ 2.16GHz", 1182);
            dict.Add("Intel Core2 Duo T5870 @ 2.00GHz", 1144);
            dict.Add("Intel Core2 Duo T5900 @ 2.20GHz", 1212);
            dict.Add("Intel Core2 Duo T6400 @ 2.00GHz", 1228);
            dict.Add("Intel Core2 Duo T6500 @ 2.10GHz", 1288);
            dict.Add("Intel Core2 Duo T6570 @ 2.10GHz", 1267);
            dict.Add("Intel Core2 Duo T6600 @ 2.20GHz", 1374);
            dict.Add("Intel Core2 Duo T6670 @ 2.20GHz", 1364);
            dict.Add("Intel Core2 Duo T7100 @ 1.80GHz", 1010);
            dict.Add("Intel Core2 Duo T7200 @ 2.00GHz", 1164);
            dict.Add("Intel Core2 Duo T7250 @ 2.00GHz", 1111);
            dict.Add("Intel Core2 Duo T7300 @ 2.00GHz", 1195);
            dict.Add("Intel Core2 Duo T7400 @ 2.16GHz", 1235);
            dict.Add("Intel Core2 Duo T7500 @ 2.20GHz", 1274);
            dict.Add("Intel Core2 Duo T7600 @ 2.33GHz", 1344);
            dict.Add("Intel Core2 Duo T7700 @ 2.40GHz", 1420);
            dict.Add("Intel Core2 Duo T7800 @ 2.60GHz", 1624);
            dict.Add("Intel Core2 Duo T8100 @ 2.10GHz", 1292);
            dict.Add("Intel Core2 Duo T8300 @ 2.40GHz", 1480);
            dict.Add("Intel Core2 Duo T9300 @ 2.50GHz", 1665);
            dict.Add("Intel Core2 Duo T9400 @ 2.53GHz", 1738);
            dict.Add("Intel Core2 Duo T9500 @ 2.60GHz", 1822);
            dict.Add("Intel Core2 Duo T9550 @ 2.66GHz", 1807);
            dict.Add("Intel Core2 Duo T9600 @ 2.80GHz", 1928);
            dict.Add("Intel Core2 Duo T9800 @ 2.93GHz", 2042);
            dict.Add("Intel Core2 Duo T9900 @ 3.06GHz", 2149);
            dict.Add("Intel Core2 Duo U7300 @ 1.30GHz", 876);
            dict.Add("Intel Core2 Duo U7500 @ 1.06GHz", 579);
            dict.Add("Intel Core2 Duo U7600 @ 1.20GHz", 667);
            dict.Add("Intel Core2 Duo U7700 @ 1.33GHz", 711);
            dict.Add("Intel Core2 Duo U9300 @ 1.20GHz", 861);
            dict.Add("Intel Core2 Duo U9600 @ 1.60GHz", 1135);
            dict.Add("Intel Core2 Duo X7360 @ 2.53GHz", 1687);
            dict.Add("Intel Core2 Extreme Q6800 @ 2.93GHz", 3646);
            dict.Add("Intel Core2 Extreme Q6850 @ 3.00GHz", 3699);
            dict.Add("Intel Core2 Extreme Q9300 @ 2.53GHz", 3538);
            dict.Add("Intel Core2 Extreme X6800 @ 2.93GHz", 1887);
            dict.Add("Intel Core2 Extreme X7800 @ 2.60GHz", 1802);
            dict.Add("Intel Core2 Extreme X7850 @ 2.80GHz", 1331);
            dict.Add("Intel Core2 Extreme X7900 @ 2.80GHz", 1679);
            dict.Add("Intel Core2 Extreme X9000 @ 2.80GHz", 1917);
            dict.Add("Intel Core2 Extreme X9100 @ 3.06GHz", 2047);
            dict.Add("Intel Core2 Extreme X9650 @ 3.00GHz", 4229);
            dict.Add("Intel Core2 Extreme X9750 @ 3.16GHz", 4651);
            dict.Add("Intel Core2 Extreme X9770 @ 3.20GHz", 4642);
            dict.Add("Intel Core2 Extreme X9775 @ 3.20GHz", 4646);
            dict.Add("Intel Core2 Quad Q6600 @ 2.40GHz", 2973);
            dict.Add("Intel Core2 Quad Q6700 @ 2.66GHz", 3308);
            dict.Add("Intel Core2 Quad Q8200 @ 2.33GHz", 2827);
            dict.Add("Intel Core2 Quad Q8300 @ 2.50GHz", 3006);
            dict.Add("Intel Core2 Quad Q8400 @ 2.66GHz", 3202);
            dict.Add("Intel Core2 Quad Q9000 @ 2.00GHz", 2532);
            dict.Add("Intel Core2 Quad Q9100 @ 2.26GHz", 3233);
            dict.Add("Intel Core2 Quad Q9300 @ 2.50GHz", 3176);
            dict.Add("Intel Core2 Quad Q9400 @ 2.66GHz", 3384);
            dict.Add("Intel Core2 Quad Q9450 @ 2.66GHz", 3787);
            dict.Add("Intel Core2 Quad Q9500 @ 2.83GHz", 3636);
            dict.Add("Intel Core2 Quad Q9505 @ 2.83GHz", 3615);
            dict.Add("Intel Core2 Quad Q9550 @ 2.83GHz", 4011);
            dict.Add("Intel Core2 Quad Q9650 @ 3.00GHz", 4236);
            dict.Add("Intel Core2 Quad Q9705 @ 3.16GHz", 3947);
            dict.Add("Intel Core2 Solo U2100 @ 1.06GHz", 303);
            dict.Add("Intel Core2 Solo U2200 @ 1.20GHz", 311);
            dict.Add("Intel Core2 Solo U3300 @ 1.20GHz", 415);
            dict.Add("Intel Core2 Solo U3500 @ 1.40GHz", 467);
            dict.Add("Intel Core 330 @ 1.60GHz", 643);
            dict.Add("Intel Core 860 @ 2.80GHz", 4550);
            dict.Add("Intel Core Duo L2300 @ 1.50GHz", 719);
            dict.Add("Intel Core Duo L2400 @ 1.66GHz", 705);
            dict.Add("Intel Core Duo L2500 @ 1.83GHz", 728);
            dict.Add("Intel Core Duo T2050 @ 1.60GHz", 704);
            dict.Add("Intel Core Duo T2250 @ 1.73GHz", 762);
            dict.Add("Intel Core Duo T2300 @ 1.66GHz", 733);
            dict.Add("Intel Core Duo T2350 @ 1.86GHz", 762);
            dict.Add("Intel Core Duo T2400 @ 1.83GHz", 793);
            dict.Add("Intel Core Duo T2450 @ 2.00GHz", 841);
            dict.Add("Intel Core Duo T2500 @ 2.00GHz", 876);
            dict.Add("Intel Core Duo T2600 @ 2.16GHz", 941);
            dict.Add("Intel Core Duo T2700 @ 2.33GHz", 1024);
            dict.Add("Intel Core Duo U2400 @ 1.06GHz", 544);
            dict.Add("Intel Core Duo U2500 @ 1.20GHz", 515);
            dict.Add("Intel Core i3-330E @ 2.13GHz", 1935);
            dict.Add("Intel Core i3-330M @ 2.13GHz", 1792);
            dict.Add("Intel Core i3-330UM @ 1.20GHz", 1053);
            dict.Add("Intel Core i3-350M @ 2.27GHz", 1901);
            dict.Add("Intel Core i3-370M @ 2.40GHz", 2023);
            dict.Add("Intel Core i3-380M @ 2.53GHz", 2103);
            dict.Add("Intel Core i3-380UM @ 1.33GHz", 1162);
            dict.Add("Intel Core i3-390M @ 2.67GHz", 2173);
            dict.Add("Intel Core i3-530 @ 2.93GHz", 2586);
            dict.Add("Intel Core i3-540 @ 3.07GHz", 2693);
            dict.Add("Intel Core i3-550 @ 3.20GHz", 2833);
            dict.Add("Intel Core i3-560 @ 3.33GHz", 2962);
            dict.Add("Intel Core i3-2100 @ 3.10GHz", 3662);
            dict.Add("Intel Core i3-2100T @ 2.50GHz", 2898);
            dict.Add("Intel Core i3-2102 @ 3.10GHz", 3817);
            dict.Add("Intel Core i3-2105 @ 3.10GHz", 3732);
            dict.Add("Intel Core i3-2120 @ 3.30GHz", 3896);
            dict.Add("Intel Core i3-2120T @ 2.60GHz", 3155);
            dict.Add("Intel Core i3-2125 @ 3.30GHz", 4006);
            dict.Add("Intel Core i3-2130 @ 3.40GHz", 4050);
            dict.Add("Intel Core i3-2140 @ 3.50GHz", 4325);
            dict.Add("Intel Core i3-2310E @ 2.10GHz", 2839);
            dict.Add("Intel Core i3-2310M @ 2.10GHz", 2420);
            dict.Add("Intel Core i3-2312M @ 2.10GHz", 2376);
            dict.Add("Intel Core i3-2328M @ 2.20GHz", 2517);
            dict.Add("Intel Core i3-2330E @ 2.20GHz", 3073);
            dict.Add("Intel Core i3-2330M @ 2.20GHz", 2528);
            dict.Add("Intel Core i3-2332M @ 2.20GHz", 2382);
            dict.Add("Intel Core i3-2340UE @ 1.30GHz", 1719);
            dict.Add("Intel Core i3-2348M @ 2.30GHz", 2604);
            dict.Add("Intel Core i3-2350M @ 2.30GHz", 2613);
            dict.Add("Intel Core i3-2357M @ 1.30GHz", 1580);
            dict.Add("Intel Core i3-2365M @ 1.40GHz", 1672);
            dict.Add("Intel Core i3-2367M @ 1.40GHz", 1696);
            dict.Add("Intel Core i3-2370M @ 2.40GHz", 2752);
            dict.Add("Intel Core i3-2375M @ 1.50GHz", 1776);
            dict.Add("Intel Core i3-2377M @ 1.50GHz", 1822);
            dict.Add("Intel Core i3-3110M @ 2.40GHz", 3064);
            dict.Add("Intel Core i3-3120M @ 2.50GHz", 3215);
            dict.Add("Intel Core i3-3130M @ 2.60GHz", 3362);
            dict.Add("Intel Core i3-3210 @ 3.20GHz", 4020);
            dict.Add("Intel Core i3-3217U @ 1.80GHz", 2298);
            dict.Add("Intel Core i3-3217UE @ 1.60GHz", 2162);
            dict.Add("Intel Core i3-3220 @ 3.30GHz", 4221);
            dict.Add("Intel Core i3-3220T @ 2.80GHz", 3698);
            dict.Add("Intel Core i3-3225 @ 3.30GHz", 4337);
            dict.Add("Intel Core i3-3227U @ 1.90GHz", 2450);
            dict.Add("Intel Core i3-3229Y @ 1.40GHz", 1676);
            dict.Add("Intel Core i3-3240 @ 3.40GHz", 4304);
            dict.Add("Intel Core i3-3240T @ 2.90GHz", 3668);
            dict.Add("Intel Core i3-3245 @ 3.40GHz", 4382);
            dict.Add("Intel Core i3-3250 @ 3.50GHz", 4429);
            dict.Add("Intel Core i3-3250T @ 3.00GHz", 3021);
            dict.Add("Intel Core i3-4000M @ 2.40GHz", 3235);
            dict.Add("Intel Core i3-4005U @ 1.70GHz", 2455);
            dict.Add("Intel Core i3-4010U @ 1.70GHz", 2448);
            dict.Add("Intel Core i3-4010Y @ 1.30GHz", 1852);
            dict.Add("Intel Core i3-4012Y @ 1.50GHz", 2199);
            dict.Add("Intel Core i3-4020Y @ 1.50GHz", 2192);
            dict.Add("Intel Core i3-4025U @ 1.90GHz", 2794);
            dict.Add("Intel Core i3-4030U @ 1.90GHz", 2697);
            dict.Add("Intel Core i3-4030Y @ 1.60GHz", 2471);
            dict.Add("Intel Core i3-4100M @ 2.50GHz", 3443);
            dict.Add("Intel Core i3-4110M @ 2.60GHz", 3873);
            dict.Add("Intel Core i3-4110U @ 1.90GHz", 2902);
            dict.Add("Intel Core i3-4120U @ 2.00GHz", 3050);
            dict.Add("Intel Core i3-4130 @ 3.40GHz", 4779);
            dict.Add("Intel Core i3-4130T @ 2.90GHz", 4097);
            dict.Add("Intel Core i3-4150 @ 3.50GHz", 4909);
            dict.Add("Intel Core i3-4150T @ 3.00GHz", 4204);
            dict.Add("Intel Core i3-4158U @ 2.00GHz", 2914);
            dict.Add("Intel Core i3-4160 @ 3.60GHz", 5036);
            dict.Add("Intel Core i3-4160T @ 3.10GHz", 4370);
            dict.Add("Intel Core i3-4170 @ 3.70GHz", 5167);
            dict.Add("Intel Core i3-4170T @ 3.20GHz", 4550);
            dict.Add("Intel Core i3-4330 @ 3.50GHz", 5073);
            dict.Add("Intel Core i3-4330T @ 3.00GHz", 4509);
            dict.Add("Intel Core i3-4330TE @ 2.40GHz", 3251);
            dict.Add("Intel Core i3-4340 @ 3.60GHz", 5230);
            dict.Add("Intel Core i3-4350 @ 3.60GHz", 4888);
            dict.Add("Intel Core i3-4350T @ 3.10GHz", 4302);
            dict.Add("Intel Core i3-4360 @ 3.70GHz", 5462);
            dict.Add("Intel Core i3-4360T @ 3.20GHz", 4643);
            dict.Add("Intel Core i3-4370 @ 3.80GHz", 5579);
            dict.Add("Intel Core i3-4570T @ 2.90GHz", 4934);
            dict.Add("Intel Core i3-5005U @ 2.00GHz", 2919);
            dict.Add("Intel Core i3-5010U @ 2.10GHz", 3062);
            dict.Add("Intel Core i3-5015U @ 2.10GHz", 3065);
            dict.Add("Intel Core i3-5020U @ 2.20GHz", 3206);
            dict.Add("Intel Core i3-5157U @ 2.50GHz", 3638);
            dict.Add("Intel Core i3-6006U @ 2.00GHz", 3164);
            dict.Add("Intel Core i3-6098P @ 3.60GHz", 6006);
            dict.Add("Intel Core i3-6100 @ 3.70GHz", 5476);
            dict.Add("Intel Core i3-6100E @ 2.70GHz", 3769);
            dict.Add("Intel Core i3-6100H @ 2.70GHz", 4118);
            dict.Add("Intel Core i3-6100T @ 3.20GHz", 4846);
            dict.Add("Intel Core i3-6100TE @ 2.70GHz", 4355);
            dict.Add("Intel Core i3-6100U @ 2.30GHz", 3874);
            dict.Add("Intel Core i3-6157U @ 2.40GHz", 3807);
            dict.Add("Intel Core i3-6300 @ 3.80GHz", 5818);
            dict.Add("Intel Core i3-6300T @ 3.30GHz", 5217);
            dict.Add("Intel Core i3-6320 @ 3.90GHz", 6024);
            dict.Add("Intel Core i3-7100 @ 3.90GHz", 5954);
            dict.Add("Intel Core i3-7100T @ 3.40GHz", 5283);
            dict.Add("Intel Core i3-7100U @ 2.40GHz", 3828);
            dict.Add("Intel Core i3-7300 @ 4.00GHz", 6484);
            dict.Add("Intel Core i3-7320 @ 4.10GHz", 6491);
            dict.Add("Intel Core i3-7350K @ 4.20GHz", 6782);
            dict.Add("Intel Core i5 750S @ 2.40GHz", 2197);
            dict.Add("Intel Core i5 E 520 @ 2.40GHz", 2486);
            dict.Add("Intel Core i5-7Y54 @ 1.20GHz", 3425);
            dict.Add("Intel Core i5-7Y57 @ 1.20GHz", 4529);
            dict.Add("Intel Core i5-430M @ 2.27GHz", 2107);
            dict.Add("Intel Core i5-430UM @ 1.20GHz", 1336);
            dict.Add("Intel Core i5-450M @ 2.40GHz", 2116);
            dict.Add("Intel Core i5-460M @ 2.53GHz", 2335);
            dict.Add("Intel Core i5-470UM @ 1.33GHz", 1274);
            dict.Add("Intel Core i5-480M @ 2.67GHz", 2425);
            dict.Add("Intel Core i5-520 @ 2.40GHz", 2244);
            dict.Add("Intel Core i5-520M @ 2.40GHz", 2382);
            dict.Add("Intel Core i5-520UM @ 1.07GHz", 1349);
            dict.Add("Intel Core i5-540M @ 2.53GHz", 2448);
            dict.Add("Intel Core i5-540UM @ 1.20GHz", 1389);
            dict.Add("Intel Core i5-560M @ 2.67GHz", 2601);
            dict.Add("Intel Core i5-560UM @ 1.33GHz", 1727);
            dict.Add("Intel Core i5-580M @ 2.67GHz", 2643);
            dict.Add("Intel Core i5-650 @ 3.20GHz", 3116);
            dict.Add("Intel Core i5-655K @ 3.20GHz", 3299);
            dict.Add("Intel Core i5-660 @ 3.33GHz", 3283);
            dict.Add("Intel Core i5-661 @ 3.33GHz", 3184);
            dict.Add("Intel Core i5-670 @ 3.47GHz", 3310);
            dict.Add("Intel Core i5-680 @ 3.60GHz", 3502);
            dict.Add("Intel Core i5-750 @ 2.67GHz", 3715);
            dict.Add("Intel Core i5-760 @ 2.80GHz", 3906);
            dict.Add("Intel Core i5-2300 @ 2.80GHz", 5308);
            dict.Add("Intel Core i5-2310 @ 2.90GHz", 5494);
            dict.Add("Intel Core i5-2320 @ 3.00GHz", 5725);
            dict.Add("Intel Core i5-2380P @ 3.10GHz", 5655);
            dict.Add("Intel Core i5-2390T @ 2.70GHz", 4276);
            dict.Add("Intel Core i5-2400 @ 3.10GHz", 5888);
            dict.Add("Intel Core i5-2400S @ 2.50GHz", 4908);
            dict.Add("Intel Core i5-2405S @ 2.50GHz", 5004);
            dict.Add("Intel Core i5-2410M @ 2.30GHz", 3158);
            dict.Add("Intel Core i5-2415M @ 2.30GHz", 3298);
            dict.Add("Intel Core i5-2430M @ 2.40GHz", 3276);
            dict.Add("Intel Core i5-2435M @ 2.40GHz", 3287);
            dict.Add("Intel Core i5-2450M @ 2.50GHz", 3404);
            dict.Add("Intel Core i5-2450P @ 3.20GHz", 6119);
            dict.Add("Intel Core i5-2467M @ 1.60GHz", 2322);
            dict.Add("Intel Core i5-2500 @ 3.30GHz", 6275);
            dict.Add("Intel Core i5-2500K @ 3.30GHz", 6460);
            dict.Add("Intel Core i5-2500S @ 2.70GHz", 5258);
            dict.Add("Intel Core i5-2500T @ 2.30GHz", 4680);
            dict.Add("Intel Core i5-2510E @ 2.50GHz", 3614);
            dict.Add("Intel Core i5-2515E @ 2.50GHz", 3313);
            dict.Add("Intel Core i5-2520M @ 2.50GHz", 3566);
            dict.Add("Intel Core i5-2537M @ 1.40GHz", 2105);
            dict.Add("Intel Core i5-2540M @ 2.60GHz", 3760);
            dict.Add("Intel Core i5-2550K @ 3.40GHz", 6699);
            dict.Add("Intel Core i5-2557M @ 1.70GHz", 2667);
            dict.Add("Intel Core i5-2560M @ 2.70GHz", 3752);
            dict.Add("Intel Core i5-3210M @ 2.50GHz", 3800);
            dict.Add("Intel Core i5-3230M @ 2.60GHz", 3922);
            dict.Add("Intel Core i5-3317U @ 1.70GHz", 3090);
            dict.Add("Intel Core i5-3320M @ 2.60GHz", 4060);
            dict.Add("Intel Core i5-3330 @ 3.00GHz", 5892);
            dict.Add("Intel Core i5-3330S @ 2.70GHz", 5621);
            dict.Add("Intel Core i5-3335S @ 2.70GHz", 5768);
            dict.Add("Intel Core i5-3337U @ 1.80GHz", 3211);
            dict.Add("Intel Core i5-3339Y @ 1.50GHz", 2228);
            dict.Add("Intel Core i5-3340 @ 3.10GHz", 6040);
            dict.Add("Intel Core i5-3340M @ 2.70GHz", 4223);
            dict.Add("Intel Core i5-3340S @ 2.80GHz", 5730);
            dict.Add("Intel Core i5-3350P @ 3.10GHz", 6106);
            dict.Add("Intel Core i5-3360M @ 2.80GHz", 4384);
            dict.Add("Intel Core i5-3380M @ 2.90GHz", 4416);
            dict.Add("Intel Core i5-3427U @ 1.80GHz", 3534);
            dict.Add("Intel Core i5-3437U @ 1.90GHz", 3611);
            dict.Add("Intel Core i5-3439Y @ 1.50GHz", 2957);
            dict.Add("Intel Core i5-3450 @ 3.10GHz", 6472);
            dict.Add("Intel Core i5-3450S @ 2.80GHz", 6129);
            dict.Add("Intel Core i5-3470 @ 3.20GHz", 6610);
            dict.Add("Intel Core i5-3470S @ 2.90GHz", 6255);
            dict.Add("Intel Core i5-3470T @ 2.90GHz", 4493);
            dict.Add("Intel Core i5-3475S @ 2.90GHz", 6398);
            dict.Add("Intel Core i5-3550 @ 3.30GHz", 6856);
            dict.Add("Intel Core i5-3550S @ 3.00GHz", 6834);
            dict.Add("Intel Core i5-3570 @ 3.40GHz", 7022);
            dict.Add("Intel Core i5-3570K @ 3.40GHz", 7144);
            dict.Add("Intel Core i5-3570S @ 3.10GHz", 6636);
            dict.Add("Intel Core i5-3570T @ 2.30GHz", 5877);
            dict.Add("Intel Core i5-3610ME @ 2.70GHz", 3722);
            dict.Add("Intel Core i5-4200H @ 2.80GHz", 4409);
            dict.Add("Intel Core i5-4200M @ 2.50GHz", 4033);
            dict.Add("Intel Core i5-4200U @ 1.60GHz", 3269);
            dict.Add("Intel Core i5-4200Y @ 1.40GHz", 2396);
            dict.Add("Intel Core i5-4202Y @ 1.60GHz", 2262);
            dict.Add("Intel Core i5-4210H @ 2.90GHz", 4509);
            dict.Add("Intel Core i5-4210M @ 2.60GHz", 4199);
            dict.Add("Intel Core i5-4210U @ 1.70GHz", 3384);
            dict.Add("Intel Core i5-4210Y @ 1.50GHz", 2359);
            dict.Add("Intel Core i5-4220Y @ 1.60GHz", 2262);
            dict.Add("Intel Core i5-4250U @ 1.30GHz", 3433);
            dict.Add("Intel Core i5-4258U @ 2.40GHz", 4079);
            dict.Add("Intel Core i5-4260U @ 1.40GHz", 3543);
            dict.Add("Intel Core i5-4278U @ 2.60GHz", 4323);
            dict.Add("Intel Core i5-4288U @ 2.60GHz", 4433);
            dict.Add("Intel Core i5-4300M @ 2.60GHz", 4379);
            dict.Add("Intel Core i5-4300U @ 1.90GHz", 3741);
            dict.Add("Intel Core i5-4300Y @ 1.60GHz", 2527);
            dict.Add("Intel Core i5-4302Y @ 1.60GHz", 3013);
            dict.Add("Intel Core i5-4308U @ 2.80GHz", 4358);
            dict.Add("Intel Core i5-4310M @ 2.70GHz", 4530);
            dict.Add("Intel Core i5-4310U @ 2.00GHz", 3675);
            dict.Add("Intel Core i5-4330M @ 2.80GHz", 4509);
            dict.Add("Intel Core i5-4340M @ 2.90GHz", 4764);
            dict.Add("Intel Core i5-4350U @ 1.40GHz", 3610);
            dict.Add("Intel Core i5-4400E @ 2.70GHz", 1271);
            dict.Add("Intel Core i5-4402E @ 1.60GHz", 3947);
            dict.Add("Intel Core i5-4430 @ 3.00GHz", 6272);
            dict.Add("Intel Core i5-4430S @ 2.70GHz", 5886);
            dict.Add("Intel Core i5-4440 @ 3.10GHz", 6445);
            dict.Add("Intel Core i5-4440S @ 2.80GHz", 6115);
            dict.Add("Intel Core i5-4460 @ 3.20GHz", 6654);
            dict.Add("Intel Core i5-4460S @ 2.90GHz", 6548);
            dict.Add("Intel Core i5-4460T @ 1.90GHz", 4928);
            dict.Add("Intel Core i5-4570 @ 3.20GHz", 7051);
            dict.Add("Intel Core i5-4570R @ 2.70GHz", 6632);
            dict.Add("Intel Core i5-4570S @ 2.90GHz", 6675);
            dict.Add("Intel Core i5-4570T @ 2.90GHz", 4794);
            dict.Add("Intel Core i5-4570TE @ 2.70GHz", 3861);
            dict.Add("Intel Core i5-4590 @ 3.30GHz", 7222);
            dict.Add("Intel Core i5-4590S @ 3.00GHz", 6949);
            dict.Add("Intel Core i5-4590T @ 2.00GHz", 5500);
            dict.Add("Intel Core i5-4670 @ 3.40GHz", 7372);
            dict.Add("Intel Core i5-4670K @ 3.40GHz", 7614);
            dict.Add("Intel Core i5-4670K CPT @ 3.40GHz", 7411);
            dict.Add("Intel Core i5-4670S @ 3.10GHz", 6436);
            dict.Add("Intel Core i5-4670T @ 2.30GHz", 6229);
            dict.Add("Intel Core i5-4690 @ 3.50GHz", 7596);
            dict.Add("Intel Core i5-4690K @ 3.50GHz", 7756);
            dict.Add("Intel Core i5-4690S @ 3.20GHz", 7360);
            dict.Add("Intel Core i5-4690T @ 2.50GHz", 6410);
            dict.Add("Intel Core i5-5200U @ 2.20GHz", 3504);
            dict.Add("Intel Core i5-5250U @ 1.60GHz", 3605);
            dict.Add("Intel Core i5-5257U @ 2.70GHz", 4375);
            dict.Add("Intel Core i5-5287U @ 2.90GHz", 4710);
            dict.Add("Intel Core i5-5300U @ 2.30GHz", 3753);
            dict.Add("Intel Core i5-5350U @ 1.80GHz", 2493);
            dict.Add("Intel Core i5-5575R @ 2.80GHz", 7278);
            dict.Add("Intel Core i5-5675C @ 3.10GHz", 8101);
            dict.Add("Intel Core i5-5675R @ 3.10GHz", 7687);
            dict.Add("Intel Core i5-6198DU @ 2.30GHz", 4195);
            dict.Add("Intel Core i5-6200U @ 2.30GHz", 3969);
            dict.Add("Intel Core i5-6260U @ 1.80GHz", 4381);
            dict.Add("Intel Core i5-6267U @ 2.90GHz", 5014);
            dict.Add("Intel Core i5-6300HQ @ 2.30GHz", 6017);
            dict.Add("Intel Core i5-6300U @ 2.40GHz", 4368);
            dict.Add("Intel Core i5-6360U @ 2.00GHz", 5038);
            dict.Add("Intel Core i5-6400 @ 2.70GHz", 6711);
            dict.Add("Intel Core i5-6400T @ 2.20GHz", 5546);
            dict.Add("Intel Core i5-6402P @ 2.80GHz", 7764);
            dict.Add("Intel Core i5-6440EQ @ 2.70GHz", 5778);
            dict.Add("Intel Core i5-6440HQ @ 2.60GHz", 6675);
            dict.Add("Intel Core i5-6500 @ 3.20GHz", 7212);
            dict.Add("Intel Core i5-6500T @ 2.50GHz", 6176);
            dict.Add("Intel Core i5-6500TE @ 2.30GHz", 6437);
            dict.Add("Intel Core i5-6600 @ 3.30GHz", 7831);
            dict.Add("Intel Core i5-6600K @ 3.50GHz", 7996);
            dict.Add("Intel Core i5-6600T @ 2.70GHz", 7244);
            dict.Add("Intel Core i5-7200U @ 2.50GHz", 4711);
            dict.Add("Intel Core i5-7260U @ 2.20GHz", 5846);
            dict.Add("Intel Core i5-7300HQ @ 2.50GHz", 6637);
            dict.Add("Intel Core i5-7300U @ 2.60GHz", 5136);
            dict.Add("Intel Core i5-7400 @ 3.00GHz", 7450);
            dict.Add("Intel Core i5-7400T @ 2.40GHz", 6450);
            dict.Add("Intel Core i5-7440HQ @ 2.80GHz", 7701);
            dict.Add("Intel Core i5-7500 @ 3.40GHz", 7971);
            dict.Add("Intel Core i5-7500T @ 2.70GHz", 7014);
            dict.Add("Intel Core i5-7600 @ 3.50GHz", 8972);
            dict.Add("Intel Core i5-7600K @ 3.80GHz", 9313);
            dict.Add("Intel Core i5-7600T @ 2.80GHz", 8455);
            dict.Add("Intel Core i5-24050S @ 2.50GHz", 4928);
            dict.Add("Intel Core i7-7Y75 @ 1.30GHz", 3965);
            dict.Add("Intel Core i7-610 @ 2.53GHz", 1969);
            dict.Add("Intel Core i7-610E @ 2.53GHz", 2544);
            dict.Add("Intel Core i7-620LM @ 2.00GHz", 1967);
            dict.Add("Intel Core i7-620M @ 2.67GHz", 2751);
            dict.Add("Intel Core i7-620UM @ 1.07GHz", 1323);
            dict.Add("Intel Core i7-640LM @ 2.13GHz", 2238);
            dict.Add("Intel Core i7-640M @ 2.80GHz", 2872);
            dict.Add("Intel Core i7-640UM @ 1.20GHz", 1688);
            dict.Add("Intel Core i7-660UM @ 1.33GHz", 1900);
            dict.Add("Intel Core i7-680UM @ 1.47GHz", 1686);
            dict.Add("Intel Core i7-720QM @ 1.60GHz", 3033);
            dict.Add("Intel Core i7-740QM @ 1.73GHz", 3210);
            dict.Add("Intel Core i7-820QM @ 1.73GHz", 3243);
            dict.Add("Intel Core i7-840QM @ 1.87GHz", 3420);
            dict.Add("Intel Core i7-860 @ 2.80GHz", 5060);
            dict.Add("Intel Core i7-860S @ 2.53GHz", 4776);
            dict.Add("Intel Core i7-870 @ 2.93GHz", 5406);
            dict.Add("Intel Core i7-870S @ 2.67GHz", 4969);
            dict.Add("Intel Core i7-875K @ 2.93GHz", 5466);
            dict.Add("Intel Core i7-880 @ 3.07GHz", 5686);
            dict.Add("Intel Core i7-920 @ 2.67GHz", 4967);
            dict.Add("Intel Core i7-920XM @ 2.00GHz", 3802);
            dict.Add("Intel Core i7-930 @ 2.80GHz", 5182);
            dict.Add("Intel Core i7-940 @ 2.93GHz", 5419);
            dict.Add("Intel Core i7-940XM @ 2.13GHz", 3975);
            dict.Add("Intel Core i7-950 @ 3.07GHz", 5601);
            dict.Add("Intel Core i7-960 @ 3.20GHz", 5878);
            dict.Add("Intel Core i7-965 @ 3.20GHz", 5892);
            dict.Add("Intel Core i7-970 @ 3.20GHz", 8450);
            dict.Add("Intel Core i7-975 @ 3.33GHz", 6186);
            dict.Add("Intel Core i7-980 @ 3.33GHz", 8810);
            dict.Add("Intel Core i7-980X @ 3.33GHz", 8897);
            dict.Add("Intel Core i7-985 @ 3.47GHz", 6641);
            dict.Add("Intel Core i7-990X @ 3.47GHz", 9164);
            dict.Add("Intel Core i7-995X @ 3.60GHz", 9791);
            dict.Add("Intel Core i7-2600 @ 3.40GHz", 8224);
            dict.Add("Intel Core i7-2600K @ 3.40GHz", 8486);
            dict.Add("Intel Core i7-2600S @ 2.80GHz", 7076);
            dict.Add("Intel Core i7-2610UE @ 1.50GHz", 2481);
            dict.Add("Intel Core i7-2617M @ 1.50GHz", 2794);
            dict.Add("Intel Core i7-2620M @ 2.70GHz", 3822);
            dict.Add("Intel Core i7-2630QM @ 2.00GHz", 5546);
            dict.Add("Intel Core i7-2630UM @ 1.60GHz", 3204);
            dict.Add("Intel Core i7-2635QM @ 2.00GHz", 5494);
            dict.Add("Intel Core i7-2637M @ 1.70GHz", 2908);
            dict.Add("Intel Core i7-2640M @ 2.80GHz", 3932);
            dict.Add("Intel Core i7-2655LE @ 2.20GHz", 3057);
            dict.Add("Intel Core i7-2670QM @ 2.20GHz", 5926);
            dict.Add("Intel Core i7-2675QM @ 2.20GHz", 5537);
            dict.Add("Intel Core i7-2677M @ 1.80GHz", 2839);
            dict.Add("Intel Core i7-2700K @ 3.50GHz", 8769);
            dict.Add("Intel Core i7-2710QE @ 2.10GHz", 5616);
            dict.Add("Intel Core i7-2715QE @ 2.10GHz", 5262);
            dict.Add("Intel Core i7-2720QM @ 2.20GHz", 6111);
            dict.Add("Intel Core i7-2760QM @ 2.40GHz", 6628);
            dict.Add("Intel Core i7-2820QM @ 2.30GHz", 6637);
            dict.Add("Intel Core i7-2840QM @ 2.40GHz", 6766);
            dict.Add("Intel Core i7-2860QM @ 2.50GHz", 7049);
            dict.Add("Intel Core i7-2920XM @ 2.50GHz", 7089);
            dict.Add("Intel Core i7-2960XM @ 2.70GHz", 7219);
            dict.Add("Intel Core i7-3517U @ 1.90GHz", 3598);
            dict.Add("Intel Core i7-3517UE @ 1.70GHz", 3259);
            dict.Add("Intel Core i7-3520M @ 2.90GHz", 4518);
            dict.Add("Intel Core i7-3537U @ 2.00GHz", 3842);
            dict.Add("Intel Core i7-3540M @ 3.00GHz", 4669);
            dict.Add("Intel Core i7-3555LE @ 2.50GHz", 4130);
            dict.Add("Intel Core i7-3610QE @ 2.30GHz", 6472);
            dict.Add("Intel Core i7-3610QM @ 2.30GHz", 7463);
            dict.Add("Intel Core i7-3612QE @ 2.10GHz", 6340);
            dict.Add("Intel Core i7-3612QM @ 2.10GHz", 6823);
            dict.Add("Intel Core i7-3615QE @ 2.30GHz", 7133);
            dict.Add("Intel Core i7-3615QM @ 2.30GHz", 7380);
            dict.Add("Intel Core i7-3630QM @ 2.40GHz", 7594);
            dict.Add("Intel Core i7-3632QM @ 2.20GHz", 6931);
            dict.Add("Intel Core i7-3635QM @ 2.40GHz", 6633);
            dict.Add("Intel Core i7-3667U @ 2.00GHz", 3949);
            dict.Add("Intel Core i7-3687U @ 2.10GHz", 4244);
            dict.Add("Intel Core i7-3689Y @ 1.50GHz", 3219);
            dict.Add("Intel Core i7-3720QM @ 2.60GHz", 8150);
            dict.Add("Intel Core i7-3740QM @ 2.70GHz", 8345);
            dict.Add("Intel Core i7-3770 @ 3.40GHz", 9314);
            dict.Add("Intel Core i7-3770K @ 3.50GHz", 9549);
            dict.Add("Intel Core i7-3770S @ 3.10GHz", 8878);
            dict.Add("Intel Core i7-3770T @ 2.50GHz", 8196);
            dict.Add("Intel Core i7-3820 @ 3.60GHz", 8993);
            dict.Add("Intel Core i7-3820QM @ 2.70GHz", 8494);
            dict.Add("Intel Core i7-3840QM @ 2.80GHz", 8839);
            dict.Add("Intel Core i7-3920XM @ 2.90GHz", 9224);
            dict.Add("Intel Core i7-3930K @ 3.20GHz", 12031);
            dict.Add("Intel Core i7-3940XM @ 3.00GHz", 9331);
            dict.Add("Intel Core i7-3960X @ 3.30GHz", 12696);
            dict.Add("Intel Core i7-3970X @ 3.50GHz", 12637);
            dict.Add("Intel Core i7-4500U @ 1.80GHz", 3785);
            dict.Add("Intel Core i7-4510U @ 2.00GHz", 3930);
            dict.Add("Intel Core i7-4550U @ 1.50GHz", 3852);
            dict.Add("Intel Core i7-4558U @ 2.80GHz", 4317);
            dict.Add("Intel Core i7-4560U @ 1.60GHz", 4329);
            dict.Add("Intel Core i7-4578U @ 3.00GHz", 4838);
            dict.Add("Intel Core i7-4600M @ 2.90GHz", 4812);
            dict.Add("Intel Core i7-4600U @ 2.10GHz", 4106);
            dict.Add("Intel Core i7-4610M @ 3.00GHz", 4979);
            dict.Add("Intel Core i7-4610Y @ 1.70GHz", 3738);
            dict.Add("Intel Core i7-4650U @ 1.70GHz", 4024);
            dict.Add("Intel Core i7-4700EQ @ 2.40GHz", 7229);
            dict.Add("Intel Core i7-4700HQ @ 2.40GHz", 7762);
            dict.Add("Intel Core i7-4700MQ @ 2.40GHz", 7715);
            dict.Add("Intel Core i7-4702HQ @ 2.20GHz", 7530);
            dict.Add("Intel Core i7-4702MQ @ 2.20GHz", 7131);
            dict.Add("Intel Core i7-4710HQ @ 2.50GHz", 7814);
            dict.Add("Intel Core i7-4710MQ @ 2.50GHz", 7988);
            dict.Add("Intel Core i7-4712HQ @ 2.30GHz", 7494);
            dict.Add("Intel Core i7-4712MQ @ 2.30GHz", 7177);
            dict.Add("Intel Core i7-4720HQ @ 2.60GHz", 8028);
            dict.Add("Intel Core i7-4722HQ @ 2.40GHz", 8052);
            dict.Add("Intel Core i7-4750HQ @ 2.00GHz", 8274);
            dict.Add("Intel Core i7-4760HQ @ 2.10GHz", 8340);
            dict.Add("Intel Core i7-4765T @ 2.00GHz", 7300);
            dict.Add("Intel Core i7-4770 @ 3.40GHz", 9802);
            dict.Add("Intel Core i7-4770HQ @ 2.20GHz", 8935);
            dict.Add("Intel Core i7-4770K @ 3.50GHz", 10115);
            dict.Add("Intel Core i7-4770R @ 3.20GHz", 9826);
            dict.Add("Intel Core i7-4770S @ 3.10GHz", 9325);
            dict.Add("Intel Core i7-4770T @ 2.50GHz", 8675);
            dict.Add("Intel Core i7-4770TE @ 2.30GHz", 7065);
            dict.Add("Intel Core i7-4771 @ 3.50GHz", 9869);
            dict.Add("Intel Core i7-4785T @ 2.20GHz", 7568);
            dict.Add("Intel Core i7-4790 @ 3.60GHz", 10001);
            dict.Add("Intel Core i7-4790K @ 4.00GHz", 11197);
            dict.Add("Intel Core i7-4790S @ 3.20GHz", 9600);
            dict.Add("Intel Core i7-4790T @ 2.70GHz", 9062);
            dict.Add("Intel Core i7-4800MQ @ 2.70GHz", 8510);
            dict.Add("Intel Core i7-4810MQ @ 2.80GHz", 8655);
            dict.Add("Intel Core i7-4820K @ 3.70GHz", 9741);
            dict.Add("Intel Core i7-4850HQ @ 2.30GHz", 9062);
            dict.Add("Intel Core i7-4860EQ @ 1.80GHz", 7933);
            dict.Add("Intel Core i7-4860HQ @ 2.40GHz", 9297);
            dict.Add("Intel Core i7-4870HQ @ 2.50GHz", 9341);
            dict.Add("Intel Core i7-4900MQ @ 2.80GHz", 9087);
            dict.Add("Intel Core i7-4910MQ @ 2.90GHz", 9415);
            dict.Add("Intel Core i7-4930K @ 3.40GHz", 13066);
            dict.Add("Intel Core i7-4930MX @ 3.00GHz", 9521);
            dict.Add("Intel Core i7-4940MX @ 3.10GHz", 9901);
            dict.Add("Intel Core i7-4960HQ @ 2.60GHz", 9770);
            dict.Add("Intel Core i7-4960X @ 3.60GHz", 13866);
            dict.Add("Intel Core i7-4980HQ @ 2.80GHz", 10014);
            dict.Add("Intel Core i7-5500U @ 2.40GHz", 4006);
            dict.Add("Intel Core i7-5550U @ 2.00GHz", 4246);
            dict.Add("Intel Core i7-5557U @ 3.10GHz", 4930);
            dict.Add("Intel Core i7-5600U @ 2.60GHz", 4295);
            dict.Add("Intel Core i7-5650U @ 2.20GHz", 4190);
            dict.Add("Intel Core i7-5675C @ 3.10GHz", 8164);
            dict.Add("Intel Core i7-5700EQ @ 2.60GHz", 8239);
            dict.Add("Intel Core i7-5700HQ @ 2.70GHz", 8398);
            dict.Add("Intel Core i7-5775C @ 3.30GHz", 11035);
            dict.Add("Intel Core i7-5775R @ 3.30GHz", 10851);
            dict.Add("Intel Core i7-5820K @ 3.30GHz", 12994);
            dict.Add("Intel Core i7-5850HQ @ 2.70GHz", 9459);
            dict.Add("Intel Core i7-5930K @ 3.50GHz", 13642);
            dict.Add("Intel Core i7-5950HQ @ 2.90GHz", 10888);
            dict.Add("Intel Core i7-5960X @ 3.00GHz", 15974);
            dict.Add("Intel Core i7-6498DU @ 2.50GHz", 4577);
            dict.Add("Intel Core i7-6500U @ 2.50GHz", 4404);
            dict.Add("Intel Core i7-6560U @ 2.20GHz", 4825);
            dict.Add("Intel Core i7-6567U @ 3.30GHz", 5617);
            dict.Add("Intel Core i7-6600U @ 2.60GHz", 4809);
            dict.Add("Intel Core i7-6650U @ 2.20GHz", 4914);
            dict.Add("Intel Core i7-6700 @ 3.40GHz", 10038);
            dict.Add("Intel Core i7-6700HQ @ 2.60GHz", 8133);
            dict.Add("Intel Core i7-6700K @ 4.00GHz", 11118);
            dict.Add("Intel Core i7-6700T @ 2.80GHz", 9050);
            dict.Add("Intel Core i7-6700TE @ 2.40GHz", 10514);
            dict.Add("Intel Core i7-6770HQ @ 2.60GHz", 9687);
            dict.Add("Intel Core i7-6800K @ 3.40GHz", 13619);
            dict.Add("Intel Core i7-6820EQ @ 2.80GHz", 8454);
            dict.Add("Intel Core i7-6820HK @ 2.70GHz", 9085);
            dict.Add("Intel Core i7-6820HQ @ 2.70GHz", 8805);
            dict.Add("Intel Core i7-6822EQ @ 2.00GHz", 6489);
            dict.Add("Intel Core i7-6850K @ 3.60GHz", 14366);
            dict.Add("Intel Core i7-6900K @ 3.20GHz", 17833);
            dict.Add("Intel Core i7-6920HQ @ 2.90GHz", 9691);
            dict.Add("Intel Core i7-6950X @ 3.00GHz", 20022);
            dict.Add("Intel Core i7-7500U @ 2.70GHz", 5254);
            dict.Add("Intel Core i7-7560U @ 2.40GHz", 5967);
            dict.Add("Intel Core i7-7567U @ 3.50GHz", 6677);
            dict.Add("Intel Core i7-7600U @ 2.80GHz", 5581);
            dict.Add("Intel Core i7-7700 @ 3.60GHz", 10843);
            dict.Add("Intel Core i7-7700HQ @ 2.80GHz", 8981);
            dict.Add("Intel Core i7-7700K @ 4.20GHz", 12232);
            dict.Add("Intel Core i7-7700T @ 2.90GHz", 9626);
            dict.Add("Intel Core i7-7820HK @ 2.90GHz", 10208);
            dict.Add("Intel Core i7-7820HQ @ 2.90GHz", 9615);
            dict.Add("Intel Core i7-7920HQ @ 3.10GHz", 10900);
            dict.Add("Intel Core m3-6Y30 @ 0.90GHz", 3056);
            dict.Add("Intel Core m3-7Y30 @ 1.00GHz", 3653);
            dict.Add("Intel Core m5-6Y54 @ 1.10GHz", 3306);
            dict.Add("Intel Core m5-6Y57 @ 1.10GHz", 3228);
            dict.Add("Intel Core m7-6Y75 @ 1.20GHz", 3560);
            dict.Add("Intel Core M-5Y10 @ 0.80GHz", 2722);
            dict.Add("Intel Core M-5Y10a @ 0.80GHz", 2150);
            dict.Add("Intel Core M-5Y10c @ 0.80GHz", 2803);
            dict.Add("Intel Core M-5Y31 @ 0.90GHz", 2763);
            dict.Add("Intel Core M-5Y51 @ 1.10GHz", 2537);
            dict.Add("Intel Core M-5Y70 @ 1.10GHz", 3029);
            dict.Add("Intel Core M-5Y71 @ 1.20GHz", 3053);
            dict.Add("Intel Core Solo T1300 @ 1.66GHz", 395);
            dict.Add("Intel Core Solo T1350 @ 1.86GHz", 402);
            dict.Add("Intel Core Solo T1400 @ 1.83GHz", 472);
            dict.Add("Intel Core Solo U1300 @ 1.06GHz", 276);
            dict.Add("Intel Core Solo U1400 @ 1.20GHz", 282);
            dict.Add("Intel Core Solo U1500 @ 1.33GHz", 330);
            dict.Add("Intel E3000 @ 3.40GHz", 2767);
            dict.Add("Intel Pentium 4 1.40GHz", 164);
            dict.Add("Intel Pentium 4 1.50GHz", 131);
            dict.Add("Intel Pentium 4 1.60GHz", 129);
            dict.Add("Intel Pentium 4 1.70GHz", 133);
            dict.Add("Intel Pentium 4 1.80GHz", 160);
            dict.Add("Intel Pentium 4 1.90GHz", 159);
            dict.Add("Intel Pentium 4 2.00GHz", 188);
            dict.Add("Intel Pentium 4 2.20GHz", 210);
            dict.Add("Intel Pentium 4 2.26GHz", 222);
            dict.Add("Intel Pentium 4 2.40GHz", 229);
            dict.Add("Intel Pentium 4 2.50GHz", 245);
            dict.Add("Intel Pentium 4 2.53GHz", 248);
            dict.Add("Intel Pentium 4 2.60GHz", 289);
            dict.Add("Intel Pentium 4 2.66GHz", 255);
            dict.Add("Intel Pentium 4 2.80GHz", 324);
            dict.Add("Intel Pentium 4 2.93GHz", 303);
            dict.Add("Intel Pentium 4 3.00GHz", 354);
            dict.Add("Intel Pentium 4 3.06GHz", 347);
            dict.Add("Intel Pentium 4 3.20GHz", 377);
            dict.Add("Intel Pentium 4 3.40GHz", 398);
            dict.Add("Intel Pentium 4 3.46GHz", 487);
            dict.Add("Intel Pentium 4 3.60GHz", 491);
            dict.Add("Intel Pentium 4 3.73GHz", 486);
            dict.Add("Intel Pentium 4 3.80GHz", 488);
            dict.Add("Intel Pentium 4 4.00GHz", 184);
            dict.Add("Intel Pentium 4 1300MHz", 119);
            dict.Add("Intel Pentium 4 1400MHz", 127);
            dict.Add("Intel Pentium 4 1500MHz", 125);
            dict.Add("Intel Pentium 4 1600MHz", 193);
            dict.Add("Intel Pentium 4 1700MHz", 154);
            dict.Add("Intel Pentium 4 1800MHz", 162);
            dict.Add("Intel Pentium 4 Mobile 1.40GHz", 149);
            dict.Add("Intel Pentium 4 Mobile 1.50GHz", 197);
            dict.Add("Intel Pentium 4 Mobile 1.60GHz", 178);
            dict.Add("Intel Pentium 4 Mobile 1.70GHz", 208);
            dict.Add("Intel Pentium 4 Mobile 1.80GHz", 175);
            dict.Add("Intel Pentium 4 Mobile 1.90GHz", 209);
            dict.Add("Intel Pentium 4 Mobile 2.00GHz", 202);
            dict.Add("Intel Pentium 5 2.66GHz", 705);
            dict.Add("Intel Pentium 957 @ 1.20GHz", 1095);
            dict.Add("Intel Pentium 967 @ 1.30GHz", 1161);
            dict.Add("Intel Pentium 977 @ 1.40GHz", 1148);
            dict.Add("Intel Pentium 987 @ 1.50GHz", 1223);
            dict.Add("Intel Pentium 997 @ 1.60GHz", 1457);
            dict.Add("Intel Pentium 1403 @ 2.60GHz", 2747);
            dict.Add("Intel Pentium 2020M @ 2.40GHz", 2298);
            dict.Add("Intel Pentium 2030M @ 2.50GHz", 2412);
            dict.Add("Intel Pentium 2117U @ 1.80GHz", 1649);
            dict.Add("Intel Pentium 2127U @ 1.90GHz", 1833);
            dict.Add("Intel Pentium 2129Y @ 1.10GHz", 1087);
            dict.Add("Intel Pentium 3550M @ 2.30GHz", 2299);
            dict.Add("Intel Pentium 3556U @ 1.70GHz", 1710);
            dict.Add("Intel Pentium 3558U @ 1.70GHz", 1706);
            dict.Add("Intel Pentium 3560M @ 2.40GHz", 2245);
            dict.Add("Intel Pentium 3560Y @ 1.20GHz", 1266);
            dict.Add("Intel Pentium 3805U @ 1.90GHz", 1987);
            dict.Add("Intel Pentium 3825U @ 1.90GHz", 2594);
            dict.Add("Intel Pentium 4405U @ 2.10GHz", 2984);
            dict.Add("Intel Pentium 4405Y @ 1.50GHz", 2074);
            dict.Add("Intel Pentium 4415U @ 2.30GHz", 3331);
            dict.Add("Intel Pentium A1018 @ 2.10GHz", 1605);
            dict.Add("Intel Pentium B940 @ 2.00GHz", 1727);
            dict.Add("Intel Pentium B950 @ 2.10GHz", 1722);
            dict.Add("Intel Pentium B960 @ 2.20GHz", 1861);
            dict.Add("Intel Pentium B970 @ 2.30GHz", 1961);
            dict.Add("Intel Pentium B980 @ 2.40GHz", 2063);
            dict.Add("Intel Pentium D 805 @ 2.66GHz", 532);
            dict.Add("Intel Pentium D 830 @ 3.00GHz", 646);
            dict.Add("Intel Pentium D 915 @ 2.80GHz", 585);
            dict.Add("Intel Pentium D 940 @ 3.20GHz", 710);
            dict.Add("Intel Pentium D 950 @ 3.40GHz", 740);
            dict.Add("Intel Pentium D 960 @ 3.60GHz", 823);
            dict.Add("Intel Pentium E2140 @ 1.60GHz", 870);
            dict.Add("Intel Pentium E2160 @ 1.80GHz", 998);
            dict.Add("Intel Pentium E2180 @ 2.00GHz", 1088);
            dict.Add("Intel Pentium E2200 @ 2.20GHz", 1202);
            dict.Add("Intel Pentium E2210 @ 2.20GHz", 1174);
            dict.Add("Intel Pentium E2220 @ 2.40GHz", 1336);
            dict.Add("Intel Pentium E5200 @ 2.50GHz", 1490);
            dict.Add("Intel Pentium E5300 @ 2.60GHz", 1548);
            dict.Add("Intel Pentium E5400 @ 2.70GHz", 1599);
            dict.Add("Intel Pentium E5500 @ 2.80GHz", 1635);
            dict.Add("Intel Pentium E5700 @ 3.00GHz", 1743);
            dict.Add("Intel Pentium E5800 @ 3.20GHz", 1909);
            dict.Add("Intel Pentium E6300 @ 2.80GHz", 1702);
            dict.Add("Intel Pentium E6500 @ 2.93GHz", 1758);
            dict.Add("Intel Pentium E6600 @ 3.06GHz", 1886);
            dict.Add("Intel Pentium E6700 @ 3.20GHz", 1945);
            dict.Add("Intel Pentium E6800 @ 3.33GHz", 2036);
            dict.Add("Intel Pentium Extreme Edition 955 @ 3.46GHz", 905);
            dict.Add("Intel Pentium Extreme Edition 965 @ 3.73GHz", 969);
            dict.Add("Intel Pentium G620 @ 2.60GHz", 2301);
            dict.Add("Intel Pentium G620T @ 2.20GHz", 2052);
            dict.Add("Intel Pentium G630 @ 2.70GHz", 2369);
            dict.Add("Intel Pentium G630T @ 2.30GHz", 2113);
            dict.Add("Intel Pentium G640 @ 2.80GHz", 2526);
            dict.Add("Intel Pentium G640T @ 2.40GHz", 1939);
            dict.Add("Intel Pentium G645 @ 2.90GHz", 2598);
            dict.Add("Intel Pentium G645T @ 2.50GHz", 2349);
            dict.Add("Intel Pentium G840 @ 2.80GHz", 2581);
            dict.Add("Intel Pentium G850 @ 2.90GHz", 2663);
            dict.Add("Intel Pentium G860 @ 3.00GHz", 2757);
            dict.Add("Intel Pentium G870 @ 3.10GHz", 2877);
            dict.Add("Intel Pentium G2010 @ 2.80GHz", 2617);
            dict.Add("Intel Pentium G2020 @ 2.90GHz", 2774);
            dict.Add("Intel Pentium G2020T @ 2.50GHz", 2459);
            dict.Add("Intel Pentium G2030 @ 3.00GHz", 2894);
            dict.Add("Intel Pentium G2030T @ 2.60GHz", 2442);
            dict.Add("Intel Pentium G2100T @ 2.60GHz", 2802);
            dict.Add("Intel Pentium G2120 @ 3.10GHz", 3073);
            dict.Add("Intel Pentium G2130 @ 3.20GHz", 3213);
            dict.Add("Intel Pentium G2140 @ 3.30GHz", 3395);
            dict.Add("Intel Pentium G3220 @ 3.00GHz", 3156);
            dict.Add("Intel Pentium G3220T @ 2.60GHz", 2616);
            dict.Add("Intel Pentium G3240 @ 3.10GHz", 3211);
            dict.Add("Intel Pentium G3240T @ 2.70GHz", 2936);
            dict.Add("Intel Pentium G3250 @ 3.20GHz", 3261);
            dict.Add("Intel Pentium G3250T @ 2.80GHz", 3010);
            dict.Add("Intel Pentium G3258 @ 3.20GHz", 3925);
            dict.Add("Intel Pentium G3260 @ 3.30GHz", 3398);
            dict.Add("Intel Pentium G3260T @ 2.90GHz", 2900);
            dict.Add("Intel Pentium G3320TE @ 2.30GHz", 2389);
            dict.Add("Intel Pentium G3420 @ 3.20GHz", 3423);
            dict.Add("Intel Pentium G3420T @ 2.70GHz", 3008);
            dict.Add("Intel Pentium G3430 @ 3.30GHz", 3479);
            dict.Add("Intel Pentium G3440 @ 3.30GHz", 3381);
            dict.Add("Intel Pentium G3440T @ 2.80GHz", 3101);
            dict.Add("Intel Pentium G3450 @ 3.40GHz", 3731);
            dict.Add("Intel Pentium G3450T @ 2.90GHz", 2967);
            dict.Add("Intel Pentium G3460 @ 3.50GHz", 3587);
            dict.Add("Intel Pentium G3470 @ 3.60GHz", 3702);
            dict.Add("Intel Pentium G4400 @ 3.30GHz", 3601);
            dict.Add("Intel Pentium G4400T @ 2.90GHz", 3246);
            dict.Add("Intel Pentium G4500 @ 3.50GHz", 3979);
            dict.Add("Intel Pentium G4500T @ 3.00GHz", 3490);
            dict.Add("Intel Pentium G4520 @ 3.60GHz", 4196);
            dict.Add("Intel Pentium G4560 @ 3.50GHz", 5039);
            dict.Add("Intel Pentium G4560T @ 2.90GHz", 4186);
            dict.Add("Intel Pentium G4600 @ 3.60GHz", 5411);
            dict.Add("Intel Pentium G4600T @ 3.00GHz", 4456);
            dict.Add("Intel Pentium G4620 @ 3.70GHz", 5404);
            dict.Add("Intel Pentium G6950 @ 2.80GHz", 1862);
            dict.Add("Intel Pentium G6951 @ 2.80GHz", 2311);
            dict.Add("Intel Pentium G6960 @ 2.93GHz", 2112);
            dict.Add("Intel Pentium III 933S @ 933MHz", 238);
            dict.Add("Intel Pentium III 1133 @ 1133MHz", 284);
            dict.Add("Intel Pentium III 1200 @ 1200MHz", 278);
            dict.Add("Intel Pentium III 1266S @ 1266MHz", 309);
            dict.Add("Intel Pentium III 1400 @ 1400MHz", 297);
            dict.Add("Intel Pentium III 1400S @ 1400MHz", 301);
            dict.Add("Intel Pentium III Mobile 750MHz", 103);
            dict.Add("Intel Pentium III Mobile 800MHz", 182);
            dict.Add("Intel Pentium III Mobile 866MHz", 152);
            dict.Add("Intel Pentium III Mobile 933MHz", 218);
            dict.Add("Intel Pentium III Mobile 1000MHz", 245);
            dict.Add("Intel Pentium III Mobile 1066MHz", 268);
            dict.Add("Intel Pentium III Mobile 1133MHz", 251);
            dict.Add("Intel Pentium III Mobile 1200MHz", 262);
            dict.Add("Intel Pentium J2850 @ 2.41GHz", 1817);
            dict.Add("Intel Pentium J2900 @ 2.41GHz", 1977);
            dict.Add("Intel Pentium J3710 @ 1.60GHz", 2025);
            dict.Add("Intel Pentium J4205 @ 1.50GHz", 2393);
            dict.Add("Intel Pentium M 1.10GHz", 282);
            dict.Add("Intel Pentium M 1.20GHz", 298);
            dict.Add("Intel Pentium M 1.30GHz", 301);
            dict.Add("Intel Pentium M 1.40GHz", 358);
            dict.Add("Intel Pentium M 1.50GHz", 367);
            dict.Add("Intel Pentium M 1.60GHz", 368);
            dict.Add("Intel Pentium M 1.70GHz", 411);
            dict.Add("Intel Pentium M 1.73GHz", 416);
            dict.Add("Intel Pentium M 1.80GHz", 406);
            dict.Add("Intel Pentium M 1.86GHz", 438);
            dict.Add("Intel Pentium M 2.00GHz", 465);
            dict.Add("Intel Pentium M 2.10GHz", 521);
            dict.Add("Intel Pentium M 2.13GHz", 508);
            dict.Add("Intel Pentium M 2.26GHz", 520);
            dict.Add("Intel Pentium M 756 @ 1.66GHz", 484);
            dict.Add("Intel Pentium M 900MHz", 220);
            dict.Add("Intel Pentium M 1000MHz", 233);
            dict.Add("Intel Pentium M 1100MHz", 259);
            dict.Add("Intel Pentium M 1200MHz", 237);
            dict.Add("Intel Pentium M 1300MHz", 303);
            dict.Add("Intel Pentium M 1400MHz", 321);
            dict.Add("Intel Pentium M 1500MHz", 355);
            dict.Add("Intel Pentium M 1600MHz", 339);
            dict.Add("Intel Pentium M 1700MHz", 376);
            dict.Add("Intel Pentium N3510 @ 1.99GHz", 1549);
            dict.Add("Intel Pentium N3520 @ 2.16GHz", 1807);
            dict.Add("Intel Pentium N3530 @ 2.16GHz", 1894);
            dict.Add("Intel Pentium N3540 @ 2.16GHz", 1949);
            dict.Add("Intel Pentium N3700 @ 1.60GHz", 1867);
            dict.Add("Intel Pentium N3710 @ 1.60GHz", 1878);
            dict.Add("Intel Pentium N4200 @ 1.10GHz", 2054);
            dict.Add("Intel Pentium P6000 @ 1.87GHz", 1245);
            dict.Add("Intel Pentium P6100 @ 2.00GHz", 1329);
            dict.Add("Intel Pentium P6200 @ 2.13GHz", 1346);
            dict.Add("Intel Pentium P6300 @ 2.27GHz", 1384);
            dict.Add("Intel Pentium SU2700 @ 1.30GHz", 407);
            dict.Add("Intel Pentium SU4100 @ 1.30GHz", 870);
            dict.Add("Intel Pentium T1080 @ 1.73GHz", 842);
            dict.Add("Intel Pentium T2060 @ 1.60GHz", 675);
            dict.Add("Intel Pentium T2080 @ 1.73GHz", 712);
            dict.Add("Intel Pentium T2130 @ 1.86GHz", 711);
            dict.Add("Intel Pentium T2310 @ 1.46GHz", 724);
            dict.Add("Intel Pentium T2330 @ 1.60GHz", 835);
            dict.Add("Intel Pentium T2370 @ 1.73GHz", 842);
            dict.Add("Intel Pentium T2390 @ 1.86GHz", 941);
            dict.Add("Intel Pentium T2410 @ 2.00GHz", 980);
            dict.Add("Intel Pentium T3200 @ 2.00GHz", 1030);
            dict.Add("Intel Pentium T3400 @ 2.16GHz", 1112);
            dict.Add("Intel Pentium T4200 @ 2.00GHz", 1147);
            dict.Add("Intel Pentium T4300 @ 2.10GHz", 1244);
            dict.Add("Intel Pentium T4400 @ 2.20GHz", 1282);
            dict.Add("Intel Pentium T4500 @ 2.30GHz", 1339);
            dict.Add("Intel Pentium U5400 @ 1.20GHz", 838);
            dict.Add("Intel Pentium U5600 @ 1.33GHz", 977);
            dict.Add("Intel T1200 @ 1.50GHz", 435);
            dict.Add("Intel T1400 @ 1.73GHz", 935);
            dict.Add("Intel T1500 @ 1.86GHz", 965);
            dict.Add("Intel T2050 @ 2.00GHz", 959);
            dict.Add("Intel Xeon 2.00GHz", 233);
            dict.Add("Intel XEON 2.20GHz", 252);
            dict.Add("Intel Xeon 2.40GHz", 276);
            dict.Add("Intel Xeon 2.66GHz", 435);
            dict.Add("Intel Xeon 2.80GHz", 400);
            dict.Add("Intel Xeon 3.00GHz", 391);
            dict.Add("Intel Xeon 3.06GHz", 491);
            dict.Add("Intel Xeon 3.20GHz", 461);
            dict.Add("Intel Xeon 3.40GHz", 478);
            dict.Add("Intel Xeon 3.60GHz", 500);
            dict.Add("Intel Xeon 3.73GHz", 1041);
            dict.Add("Intel Xeon 3.80GHz", 604);
            dict.Add("Intel Xeon 1500MHz", 193);
            dict.Add("Intel Xeon 3040 @ 1.86GHz", 1281);
            dict.Add("Intel Xeon 3050 @ 2.13GHz", 1311);
            dict.Add("Intel Xeon 3060 @ 2.40GHz", 1518);
            dict.Add("Intel Xeon 3065 @ 2.33GHz", 1504);
            dict.Add("Intel Xeon 3070 @ 2.66GHz", 1754);
            dict.Add("Intel Xeon 3075 @ 2.66GHz", 1840);
            dict.Add("Intel Xeon 3085 @ 3.00GHz", 1801);
            dict.Add("Intel Xeon 5110 @ 1.60GHz", 1079);
            dict.Add("Intel Xeon 5120 @ 1.86GHz", 1249);
            dict.Add("Intel Xeon 5130 @ 2.00GHz", 1379);
            dict.Add("Intel Xeon 5133 @ 2.20GHz", 1598);
            dict.Add("Intel Xeon 5140 @ 2.33GHz", 1639);
            dict.Add("Intel Xeon 5148 @ 2.33GHz", 1590);
            dict.Add("Intel Xeon 5150 @ 2.66GHz", 1741);
            dict.Add("Intel Xeon 5160 @ 3.00GHz", 1973);
            dict.Add("Intel Xeon D-1518 @ 2.20GHz", 4700);
            dict.Add("Intel Xeon D-1520 @ 2.20GHz", 6396);
            dict.Add("Intel Xeon D-1521 @ 2.40GHz", 6980);
            dict.Add("Intel Xeon D-1528 @ 1.90GHz", 8612);
            dict.Add("Intel Xeon D-1540 @ 2.00GHz", 10573);
            dict.Add("Intel Xeon D-1541 @ 2.10GHz", 11333);
            dict.Add("Intel Xeon D-1567 @ 2.10GHz", 15028);
            dict.Add("Intel Xeon D-1587 @ 1.70GHz", 13489);
            dict.Add("Intel Xeon E3-1220 @ 3.10GHz", 6065);
            dict.Add("Intel Xeon E3-1220 V2 @ 3.10GHz", 6594);
            dict.Add("Intel Xeon E3-1220 v3 @ 3.10GHz", 6925);
            dict.Add("Intel Xeon E3-1220 v5 @ 3.00GHz", 7597);
            dict.Add("Intel Xeon E3-1220L @ 2.20GHz", 3563);
            dict.Add("Intel Xeon E3-1220L V2 @ 2.30GHz", 3656);
            dict.Add("Intel Xeon E3-1220L v3 @ 1.10GHz", 2110);
            dict.Add("Intel Xeon E3-1225 @ 3.10GHz", 5954);
            dict.Add("Intel Xeon E3-1225 V2 @ 3.20GHz", 6807);
            dict.Add("Intel Xeon E3-1225 v3 @ 3.20GHz", 7107);
            dict.Add("Intel Xeon E3-1225 v5 @ 3.30GHz", 7765);
            dict.Add("Intel Xeon E3-1226 v3 @ 3.30GHz", 7539);
            dict.Add("Intel Xeon E3-1230 @ 3.20GHz", 7943);
            dict.Add("Intel Xeon E3-1230 V2 @ 3.30GHz", 8849);
            dict.Add("Intel Xeon E3-1230 v3 @ 3.30GHz", 9309);
            dict.Add("Intel Xeon E3-1230 v5 @ 3.40GHz", 9708);
            dict.Add("Intel Xeon E3-1230L v3 @ 1.80GHz", 7207);
            dict.Add("Intel Xeon E3-1231 v3 @ 3.40GHz", 9634);
            dict.Add("Intel Xeon E3-1235 @ 3.20GHz", 7685);
            dict.Add("Intel Xeon E3-1235L v5 @ 2.00GHz", 6352);
            dict.Add("Intel Xeon E3-1240 @ 3.30GHz", 8001);
            dict.Add("Intel Xeon E3-1240 V2 @ 3.40GHz", 9211);
            dict.Add("Intel Xeon E3-1240 v3 @ 3.40GHz", 9704);
            dict.Add("Intel Xeon E3-1240 v5 @ 3.50GHz", 10362);
            dict.Add("Intel Xeon E3-1240L v3 @ 2.00GHz", 7508);
            dict.Add("Intel Xeon E3-1241 v3 @ 3.50GHz", 10036);
            dict.Add("Intel Xeon E3-1245 @ 3.30GHz", 8058);
            dict.Add("Intel Xeon E3-1245 V2 @ 3.40GHz", 9091);
            dict.Add("Intel Xeon E3-1245 v3 @ 3.40GHz", 9568);
            dict.Add("Intel Xeon E3-1245 v5 @ 3.50GHz", 10336);
            dict.Add("Intel Xeon E3-1246 v3 @ 3.50GHz", 9995);
            dict.Add("Intel Xeon E3-1260L @ 2.40GHz", 6534);
            dict.Add("Intel Xeon E3-1260L v5 @ 2.90GHz", 10067);
            dict.Add("Intel Xeon E3-1265L @ 2.40GHz", 6054);
            dict.Add("Intel Xeon E3-1265L V2 @ 2.50GHz", 7745);
            dict.Add("Intel Xeon E3-1265L v3 @ 2.50GHz", 8691);
            dict.Add("Intel Xeon E3-1268L v3 @ 2.30GHz", 7850);
            dict.Add("Intel Xeon E3-1270 @ 3.40GHz", 8239);
            dict.Add("Intel Xeon E3-1270 V2 @ 3.50GHz", 9469);
            dict.Add("Intel Xeon E3-1270 v3 @ 3.50GHz", 9831);
            dict.Add("Intel Xeon E3-1270 v5 @ 3.60GHz", 10186);
            dict.Add("Intel Xeon E3-1271 v3 @ 3.60GHz", 10046);
            dict.Add("Intel Xeon E3-1275 @ 3.40GHz", 8348);
            dict.Add("Intel Xeon E3-1275 V2 @ 3.50GHz", 9334);
            dict.Add("Intel Xeon E3-1275 v3 @ 3.50GHz", 9847);
            dict.Add("Intel Xeon E3-1275 v5 @ 3.60GHz", 10328);
            dict.Add("Intel Xeon E3-1275L v3 @ 2.70GHz", 8571);
            dict.Add("Intel Xeon E3-1276 v3 @ 3.60GHz", 10247);
            dict.Add("Intel Xeon E3-1280 @ 3.50GHz", 8473);
            dict.Add("Intel Xeon E3-1280 V2 @ 3.60GHz", 9746);
            dict.Add("Intel Xeon E3-1280 v3 @ 3.60GHz", 9755);
            dict.Add("Intel Xeon E3-1280 v5 @ 3.70GHz", 10502);
            dict.Add("Intel Xeon E3-1280 v6 @ 3.90GHz", 11104);
            dict.Add("Intel Xeon E3-1281 v3 @ 3.70GHz", 10193);
            dict.Add("Intel Xeon E3-1285 v3 @ 3.60GHz", 10252);
            dict.Add("Intel Xeon E3-1285L v3 @ 3.10GHz", 9984);
            dict.Add("Intel Xeon E3-1285L v4 @ 3.40GHz", 11224);
            dict.Add("Intel Xeon E3-1286 v3 @ 3.70GHz", 9388);
            dict.Add("Intel Xeon E3-1290 @ 3.60GHz", 8704);
            dict.Add("Intel Xeon E3-1290 V2 @ 3.70GHz", 9862);
            dict.Add("Intel Xeon E3-1505L v5 @ 2.00GHz", 7082);
            dict.Add("Intel Xeon E3-1505M v5 @ 2.80GHz", 8905);
            dict.Add("Intel Xeon E3-1505M v6 @ 3.00GHz", 10062);
            dict.Add("Intel Xeon E3-1515M v5 @ 2.80GHz", 10509);
            dict.Add("Intel Xeon E3-1535M v5 @ 2.90GHz", 9263);
            dict.Add("Intel Xeon E3-1535M v6 @ 3.10GHz", 11082);
            dict.Add("Intel Xeon E3-1545M v5 @ 2.90GHz", 10611);
            dict.Add("Intel Xeon E3-1575M v5 @ 3.00GHz", 10988);
            dict.Add("Intel Xeon E5-1410 @ 2.80GHz", 7312);
            dict.Add("Intel Xeon E5-1410 v2 @ 2.80GHz", 6822);
            dict.Add("Intel Xeon E5-1603 @ 2.80GHz", 5481);
            dict.Add("Intel Xeon E5-1603 v3 @ 2.80GHz", 5992);
            dict.Add("Intel Xeon E5-1603 v4 @ 2.80GHz", 5375);
            dict.Add("Intel Xeon E5-1607 @ 3.00GHz", 5785);
            dict.Add("Intel Xeon E5-1607 v2 @ 3.00GHz", 6005);
            dict.Add("Intel Xeon E5-1607 v3 @ 3.10GHz", 6897);
            dict.Add("Intel Xeon E5-1607 v4 @ 3.10GHz", 7362);
            dict.Add("Intel Xeon E5-1620 @ 3.60GHz", 9093);
            dict.Add("Intel Xeon E5-1620 v2 @ 3.70GHz", 9493);
            dict.Add("Intel Xeon E5-1620 v3 @ 3.50GHz", 9739);
            dict.Add("Intel Xeon E5-1620 v4 @ 3.50GHz", 10070);
            dict.Add("Intel Xeon E5-1630 v3 @ 3.70GHz", 10239);
            dict.Add("Intel Xeon E5-1630 v4 @ 3.70GHz", 10211);
            dict.Add("Intel Xeon E5-1650 @ 3.20GHz", 11807);
            dict.Add("Intel Xeon E5-1650 v2 @ 3.50GHz", 12629);
            dict.Add("Intel Xeon E5-1650 v3 @ 3.50GHz", 13552);
            dict.Add("Intel Xeon E5-1650 v4 @ 3.60GHz", 14277);
            dict.Add("Intel Xeon E5-1660 @ 3.30GHz", 12518);
            dict.Add("Intel Xeon E5-1660 v2 @ 3.70GHz", 13735);
            dict.Add("Intel Xeon E5-1660 v3 @ 3.00GHz", 14277);
            dict.Add("Intel Xeon E5-1660 v4 @ 3.20GHz", 16334);
            dict.Add("Intel Xeon E5-1680 v2 @ 3.00GHz", 17201);
            dict.Add("Intel Xeon E5-1680 v3 @ 3.20GHz", 16673);
            dict.Add("Intel Xeon E5-1680 v4 @ 3.40GHz", 16684);
            dict.Add("Intel Xeon E5-1681 v3 @ 2.90GHz", 18367);
            dict.Add("Intel Xeon E5-2403 @ 1.80GHz", 3489);
            dict.Add("Intel Xeon E5-2403 v2 @ 1.80GHz", 1797);
            dict.Add("Intel Xeon E5-2407 @ 2.20GHz", 3795);
            dict.Add("Intel Xeon E5-2407 v2 @ 2.40GHz", 4677);
            dict.Add("Intel Xeon E5-2418L @ 2.00GHz", 5202);
            dict.Add("Intel Xeon E5-2420 @ 1.90GHz", 7139);
            dict.Add("Intel Xeon E5-2420 v2 @ 2.20GHz", 8593);
            dict.Add("Intel Xeon E5-2430 @ 2.20GHz", 6878);
            dict.Add("Intel Xeon E5-2430 v2 @ 2.50GHz", 8608);
            dict.Add("Intel Xeon E5-2430L v2 @ 2.40GHz", 6627);
            dict.Add("Intel Xeon E5-2440 @ 2.40GHz", 9319);
            dict.Add("Intel Xeon E5-2440 v2 @ 1.90GHz", 9425);
            dict.Add("Intel Xeon E5-2450 @ 2.10GHz", 10186);
            dict.Add("Intel Xeon E5-2470 @ 2.30GHz", 11149);
            dict.Add("Intel Xeon E5-2603 @ 1.80GHz", 3518);
            dict.Add("Intel Xeon E5-2603 v2 @ 1.80GHz", 3739);
            dict.Add("Intel Xeon E5-2603 v3 @ 1.60GHz", 5087);
            dict.Add("Intel Xeon E5-2603 v4 @ 1.70GHz", 5247);
            dict.Add("Intel Xeon E5-2609 @ 2.40GHz", 4576);
            dict.Add("Intel Xeon E5-2609 v2 @ 2.50GHz", 5013);
            dict.Add("Intel Xeon E5-2609 v3 @ 1.90GHz", 5949);
            dict.Add("Intel Xeon E5-2609 v4 @ 1.70GHz", 6983);
            dict.Add("Intel Xeon E5-2618L v3 @ 2.30GHz", 12508);
            dict.Add("Intel Xeon E5-2620 @ 2.00GHz", 7971);
            dict.Add("Intel Xeon E5-2620 v2 @ 2.10GHz", 8693);
            dict.Add("Intel Xeon E5-2620 v3 @ 2.40GHz", 9986);
            dict.Add("Intel Xeon E5-2620 v4 @ 2.10GHz", 11354);
            dict.Add("Intel Xeon E5-2623 v3 @ 3.00GHz", 9097);
            dict.Add("Intel Xeon E5-2623 v4 @ 2.60GHz", 8061);
            dict.Add("Intel Xeon E5-2628L v2 @ 1.90GHz", 9405);
            dict.Add("Intel Xeon E5-2628L v3 @ 2.00GHz", 12405);
            dict.Add("Intel Xeon E5-2628L v4 @ 1.90GHz", 13041);
            dict.Add("Intel Xeon E5-2629 v3 @ 2.40GHz", 10984);
            dict.Add("Intel Xeon E5-2630 @ 2.30GHz", 8915);
            dict.Add("Intel Xeon E5-2630 v2 @ 2.60GHz", 10452);
            dict.Add("Intel Xeon E5-2630 v3 @ 2.40GHz", 12822);
            dict.Add("Intel Xeon E5-2630 v4 @ 2.20GHz", 14174);
            dict.Add("Intel Xeon E5-2630L @ 2.00GHz", 7868);
            dict.Add("Intel Xeon E5-2630L v3 @ 1.80GHz", 7767);
            dict.Add("Intel Xeon E5-2630L v4 @ 1.80GHz", 12847);
            dict.Add("Intel Xeon E5-2637 v2 @ 3.50GHz", 9452);
            dict.Add("Intel Xeon E5-2637 v3 @ 3.50GHz", 10281);
            dict.Add("Intel Xeon E5-2637 v4 @ 3.50GHz", 9858);
            dict.Add("Intel Xeon E5-2640 @ 2.50GHz", 9618);
            dict.Add("Intel Xeon E5-2640 v2 @ 2.00GHz", 9914);
            dict.Add("Intel Xeon E5-2640 v3 @ 2.60GHz", 14117);
            dict.Add("Intel Xeon E5-2640 v4 @ 2.40GHz", 14833);
            dict.Add("Intel Xeon E5-2643 @ 3.30GHz", 8467);
            dict.Add("Intel Xeon E5-2643 v2 @ 3.50GHz", 11735);
            dict.Add("Intel Xeon E5-2643 v3 @ 3.40GHz", 13852);
            dict.Add("Intel Xeon E5-2643 v4 @ 3.40GHz", 13411);
            dict.Add("Intel Xeon E5-2648L v3 @ 1.80GHz", 12332);
            dict.Add("Intel Xeon E5-2650 @ 2.00GHz", 10262);
            dict.Add("Intel Xeon E5-2650 v2 @ 2.60GHz", 13117);
            dict.Add("Intel Xeon E5-2650 v3 @ 2.30GHz", 14951);
            dict.Add("Intel Xeon E5-2650 v4 @ 2.20GHz", 15994);
            dict.Add("Intel Xeon E5-2650L @ 1.80GHz", 8676);
            dict.Add("Intel Xeon E5-2650L v3 @ 1.80GHz", 13131);
            dict.Add("Intel Xeon E5-2651 v2 @ 1.80GHz", 11176);
            dict.Add("Intel Xeon E5-2658 @ 2.10GHz", 9484);
            dict.Add("Intel Xeon E5-2658 v2 @ 2.40GHz", 14128);
            dict.Add("Intel Xeon E5-2658 v3 @ 2.20GHz", 16511);
            dict.Add("Intel Xeon E5-2658 v4 @ 2.30GHz", 16290);
            dict.Add("Intel Xeon E5-2660 @ 2.20GHz", 11188);
            dict.Add("Intel Xeon E5-2660 v2 @ 2.20GHz", 13264);
            dict.Add("Intel Xeon E5-2660 v3 @ 2.60GHz", 16161);
            dict.Add("Intel Xeon E5-2660 v4 @ 2.00GHz", 18816);
            dict.Add("Intel Xeon E5-2663 v3 @ 2.80GHz", 13802);
            dict.Add("Intel Xeon E5-2665 @ 2.40GHz", 11950);
            dict.Add("Intel Xeon E5-2667 @ 2.90GHz", 10380);
            dict.Add("Intel Xeon E5-2667 v2 @ 3.30GHz", 16512);
            dict.Add("Intel Xeon E5-2667 v3 @ 3.20GHz", 16125);
            dict.Add("Intel Xeon E5-2667 v4 @ 3.20GHz", 15916);
            dict.Add("Intel Xeon E5-2670 @ 2.60GHz", 12334);
            dict.Add("Intel Xeon E5-2670 v2 @ 2.50GHz", 14975);
            dict.Add("Intel Xeon E5-2670 v3 @ 2.30GHz", 16549);
            dict.Add("Intel Xeon E5-2673 v2 @ 3.30GHz", 16320);
            dict.Add("Intel Xeon E5-2673 v3 @ 2.40GHz", 16904);
            dict.Add("Intel Xeon E5-2673 v4 @ 2.30GHz", 21073);
            dict.Add("Intel Xeon E5-2675 v3 @ 1.80GHz", 15275);
            dict.Add("Intel Xeon E5-2676 v3 @ 2.40GHz", 17795);
            dict.Add("Intel Xeon E5-2678 v3 @ 2.50GHz", 16618);
            dict.Add("Intel Xeon E5-2679 v4 @ 2.50GHz", 25236);
            dict.Add("Intel Xeon E5-2680 @ 2.70GHz", 12931);
            dict.Add("Intel Xeon E5-2680 v2 @ 2.80GHz", 16341);
            dict.Add("Intel Xeon E5-2680 v3 @ 2.50GHz", 18761);
            dict.Add("Intel Xeon E5-2680 v4 @ 2.40GHz", 19905);
            dict.Add("Intel Xeon E5-2683 v3 @ 2.00GHz", 17504);
            dict.Add("Intel Xeon E5-2685 v3 @ 2.60GHz", 14154);
            dict.Add("Intel Xeon E5-2686 v3 @ 2.00GHz", 19255);
            dict.Add("Intel Xeon E5-2687W @ 3.10GHz", 14401);
            dict.Add("Intel Xeon E5-2687W v2 @ 3.40GHz", 16559);
            dict.Add("Intel Xeon E5-2687W v3 @ 3.10GHz", 17779);
            dict.Add("Intel Xeon E5-2687W v4 @ 3.00GHz", 20130);
            dict.Add("Intel Xeon E5-2689 @ 2.60GHz", 13747);
            dict.Add("Intel Xeon E5-2689 v4 @ 3.10GHz", 19708);
            dict.Add("Intel Xeon E5-2690 @ 2.90GHz", 14191);
            dict.Add("Intel Xeon E5-2690 v2 @ 3.00GHz", 16546);
            dict.Add("Intel Xeon E5-2690 v3 @ 2.60GHz", 19362);
            dict.Add("Intel Xeon E5-2690 v4 @ 2.60GHz", 21806);
            dict.Add("Intel Xeon E5-2692 v2 @ 2.20GHz", 16018);
            dict.Add("Intel Xeon E5-2695 v2 @ 2.40GHz", 15708);
            dict.Add("Intel Xeon E5-2695 v3 @ 2.30GHz", 20431);
            dict.Add("Intel Xeon E5-2695 v4 @ 2.10GHz", 20582);
            dict.Add("Intel Xeon E5-2696 v2 @ 2.50GHz", 16681);
            dict.Add("Intel Xeon E5-2696 v3 @ 2.30GHz", 22277);
            dict.Add("Intel Xeon E5-2696 v4 @ 2.20GHz", 22197);
            dict.Add("Intel Xeon E5-2697 v2 @ 2.70GHz", 17321);
            dict.Add("Intel Xeon E5-2697 v3 @ 2.60GHz", 21608);
            dict.Add("Intel Xeon E5-2697 v4 @ 2.30GHz", 21356);
            dict.Add("Intel Xeon E5-2698 v3 @ 2.30GHz", 21149);
            dict.Add("Intel Xeon E5-2698 v4 @ 2.20GHz", 21789);
            dict.Add("Intel Xeon E5-2699 v3 @ 2.30GHz", 22645);
            dict.Add("Intel Xeon E5-2699 v4 @ 2.20GHz", 23344);
            dict.Add("Intel Xeon E5-4603 @ 2.00GHz", 5014);
            dict.Add("Intel Xeon E5-4620 @ 2.20GHz", 8127);
            dict.Add("Intel Xeon E5-4627 v3 @ 2.60GHz", 14219);
            dict.Add("Intel Xeon E5-4627 v4 @ 2.60GHz", 15516);
            dict.Add("Intel Xeon E5-4650 @ 2.70GHz", 11960);
            dict.Add("Intel Xeon E5-4669 v4 @ 2.20GHz", 13626);
            dict.Add("Intel Xeon E3110 @ 3.00GHz", 2175);
            dict.Add("Intel Xeon E3113 @ 3.00GHz", 2427);
            dict.Add("Intel Xeon E3120 @ 3.16GHz", 2241);
            dict.Add("Intel Xeon E5205 @ 1.86GHz", 1401);
            dict.Add("Intel Xeon E5240 @ 3.00GHz", 2424);
            dict.Add("Intel Xeon E5310 @ 1.60GHz", 2264);
            dict.Add("Intel Xeon E5320 @ 1.86GHz", 2282);
            dict.Add("Intel Xeon E5335 @ 2.00GHz", 2513);
            dict.Add("Intel Xeon E5345 @ 2.33GHz", 2958);
            dict.Add("Intel Xeon E5405 @ 2.00GHz", 2897);
            dict.Add("Intel Xeon E5410 @ 2.33GHz", 3284);
            dict.Add("Intel Xeon E5420 @ 2.50GHz", 3542);
            dict.Add("Intel Xeon E5430 @ 2.66GHz", 3797);
            dict.Add("Intel Xeon E5440 @ 2.83GHz", 4009);
            dict.Add("Intel Xeon E5450 @ 3.00GHz", 4247);
            dict.Add("Intel Xeon E5462 @ 2.80GHz", 3942);
            dict.Add("Intel Xeon E5472 @ 3.00GHz", 4247);
            dict.Add("Intel Xeon E5502 @ 1.87GHz", 1381);
            dict.Add("Intel Xeon E5503 @ 2.00GHz", 1357);
            dict.Add("Intel Xeon E5504 @ 2.00GHz", 2722);
            dict.Add("Intel Xeon E5506 @ 2.13GHz", 2996);
            dict.Add("Intel Xeon E5507 @ 2.27GHz", 3148);
            dict.Add("Intel Xeon E5520 @ 2.27GHz", 4453);
            dict.Add("Intel Xeon E5530 @ 2.40GHz", 4605);
            dict.Add("Intel Xeon E5540 @ 2.53GHz", 4857);
            dict.Add("Intel Xeon E5603 @ 1.60GHz", 2362);
            dict.Add("Intel Xeon E5606 @ 2.13GHz", 3093);
            dict.Add("Intel Xeon E5607 @ 2.27GHz", 3459);
            dict.Add("Intel Xeon E5620 @ 2.40GHz", 4875);
            dict.Add("Intel Xeon E5630 @ 2.53GHz", 5166);
            dict.Add("Intel Xeon E5640 @ 2.67GHz", 5345);
            dict.Add("Intel Xeon E5645 @ 2.40GHz", 6558);
            dict.Add("Intel Xeon E5649 @ 2.53GHz", 7051);
            dict.Add("Intel Xeon E7320 @ 2.13GHz", 2413);
            dict.Add("Intel Xeon L3110 @ 3.00GHz", 2273);
            dict.Add("Intel Xeon L3360 @ 2.83GHz", 3820);
            dict.Add("Intel Xeon L3426 @ 1.87GHz", 3837);
            dict.Add("Intel Xeon L5238 @ 2.66GHz", 1823);
            dict.Add("Intel Xeon L5240 @ 3.00GHz", 2271);
            dict.Add("Intel Xeon L5310 @ 1.60GHz", 2274);
            dict.Add("Intel Xeon L5320 @ 1.86GHz", 2135);
            dict.Add("Intel Xeon L5335 @ 2.00GHz", 2756);
            dict.Add("Intel Xeon L5408 @ 2.13GHz", 2965);
            dict.Add("Intel Xeon L5410 @ 2.33GHz", 3349);
            dict.Add("Intel Xeon L5420 @ 2.50GHz", 3503);
            dict.Add("Intel Xeon L5430 @ 2.66GHz", 3818);
            dict.Add("Intel Xeon L5506 @ 2.13GHz", 3715);
            dict.Add("Intel Xeon L5520 @ 2.27GHz", 4387);
            dict.Add("Intel Xeon L5530 @ 2.40GHz", 4351);
            dict.Add("Intel Xeon L5630 @ 2.13GHz", 4420);
            dict.Add("Intel Xeon L5638 @ 2.00GHz", 5674);
            dict.Add("Intel Xeon L5639 @ 2.13GHz", 7021);
            dict.Add("Intel Xeon L5640 @ 2.27GHz", 6451);
            dict.Add("Intel Xeon L7455 @ 2.13GHz", 3836);
            dict.Add("Intel Xeon MV 3.20GHz", 683);
            dict.Add("Intel Xeon W3503 @ 2.40GHz", 1775);
            dict.Add("Intel Xeon W3505 @ 2.53GHz", 1858);
            dict.Add("Intel Xeon W3520 @ 2.67GHz", 5064);
            dict.Add("Intel Xeon W3530 @ 2.80GHz", 5367);
            dict.Add("Intel Xeon W3540 @ 2.93GHz", 5479);
            dict.Add("Intel Xeon W3550 @ 3.07GHz", 5730);
            dict.Add("Intel Xeon W3565 @ 3.20GHz", 5959);
            dict.Add("Intel Xeon W3570 @ 3.20GHz", 6165);
            dict.Add("Intel Xeon W3580 @ 3.33GHz", 6412);
            dict.Add("Intel Xeon W3670 @ 3.20GHz", 8299);
            dict.Add("Intel Xeon W3680 @ 3.33GHz", 9235);
            dict.Add("Intel Xeon W3690 @ 3.47GHz", 9450);
            dict.Add("Intel Xeon W5580 @ 3.20GHz", 5836);
            dict.Add("Intel Xeon W5590 @ 3.33GHz", 6183);
            dict.Add("Intel Xeon X3210 @ 2.13GHz", 2801);
            dict.Add("Intel Xeon X3220 @ 2.40GHz", 3108);
            dict.Add("Intel Xeon X3230 @ 2.66GHz", 3443);
            dict.Add("Intel Xeon X3320 @ 2.50GHz", 3353);
            dict.Add("Intel Xeon X3323 @ 2.50GHz", 3159);
            dict.Add("Intel Xeon X3330 @ 2.66GHz", 3571);
            dict.Add("Intel Xeon X3350 @ 2.66GHz", 3898);
            dict.Add("Intel Xeon X3353 @ 2.66GHz", 3905);
            dict.Add("Intel Xeon X3360 @ 2.83GHz", 4025);
            dict.Add("Intel Xeon X3363 @ 2.83GHz", 4074);
            dict.Add("Intel Xeon X3370 @ 3.00GHz", 4346);
            dict.Add("Intel Xeon X3380 @ 3.16GHz", 4697);
            dict.Add("Intel Xeon X3430 @ 2.40GHz", 3374);
            dict.Add("Intel Xeon X3440 @ 2.53GHz", 4577);
            dict.Add("Intel Xeon X3450 @ 2.67GHz", 4923);
            dict.Add("Intel Xeon X3460 @ 2.80GHz", 5143);
            dict.Add("Intel Xeon X3470 @ 2.93GHz", 5195);
            dict.Add("Intel Xeon X3480 @ 3.07GHz", 5732);
            dict.Add("Intel Xeon X5260 @ 3.33GHz", 2469);
            dict.Add("Intel Xeon X5270 @ 3.50GHz", 2524);
            dict.Add("Intel Xeon X5272 @ 3.40GHz", 2394);
            dict.Add("Intel Xeon X5355 @ 2.66GHz", 3266);
            dict.Add("Intel Xeon X5365 @ 3.00GHz", 3491);
            dict.Add("Intel Xeon X5450 @ 3.00GHz", 4177);
            dict.Add("Intel Xeon X5460 @ 3.16GHz", 4394);
            dict.Add("Intel Xeon X5470 @ 3.33GHz", 4670);
            dict.Add("Intel Xeon X5472 @ 3.00GHz", 4127);
            dict.Add("Intel Xeon X5482 @ 3.20GHz", 4593);
            dict.Add("Intel Xeon X5492 @ 3.40GHz", 4884);
            dict.Add("Intel Xeon X5550 @ 2.67GHz", 5416);
            dict.Add("Intel Xeon X5560 @ 2.80GHz", 5426);
            dict.Add("Intel Xeon X5570 @ 2.93GHz", 5679);
            dict.Add("Intel Xeon X5647 @ 2.93GHz", 5980);
            dict.Add("Intel Xeon X5650 @ 2.67GHz", 7527);
            dict.Add("Intel Xeon X5660 @ 2.80GHz", 7839);
            dict.Add("Intel Xeon X5667 @ 3.07GHz", 4722);
            dict.Add("Intel Xeon X5670 @ 2.93GHz", 8061);
            dict.Add("Intel Xeon X5672 @ 3.20GHz", 6507);
            dict.Add("Intel Xeon X5675 @ 3.07GHz", 8556);
            dict.Add("Intel Xeon X5677 @ 3.47GHz", 6984);
            dict.Add("Intel Xeon X5679 @ 3.20GHz", 8652);
            dict.Add("Intel Xeon X5680 @ 3.33GHz", 8767);
            dict.Add("Intel Xeon X5687 @ 3.60GHz", 7138);
            dict.Add("Intel Xeon X5690 @ 3.47GHz", 9084);
            dict.Add("Intel Xeon X5698 @ 4.40GHz", 4272);
            dict.Add("Intel Xeon X6550 @ 2.00GHz", 2977);
            dict.Add("Mobile AMD Athlon 4", 314);
            dict.Add("Mobile AMD Athlon 4 2400 +", 358);
            dict.Add("Mobile AMD Athlon 64 2700 +", 398);
            dict.Add("Mobile AMD Athlon 64 2800 +", 440);
            dict.Add("Mobile AMD Athlon 64 3000 +", 450);
            dict.Add("Mobile AMD Athlon 64 3200 +", 441);
            dict.Add("Mobile AMD Athlon 64 3400 +", 511);
            dict.Add("Mobile AMD Athlon 64 3700 +", 363);
            dict.Add("Mobile AMD Athlon 64 4000 +", 576);
            dict.Add("Mobile AMD Athlon 1400 +", 266);
            dict.Add("Mobile AMD Athlon 2500 +", 370);
            dict.Add("Mobile AMD Athlon MP-M 1800 +", 265);
            dict.Add("Mobile AMD Athlon MP-M 2000 +", 378);
            dict.Add("Mobile AMD Athlon MP-M 2400 +", 295);
            dict.Add("Mobile AMD Athlon MP-M 2800 +", 496);
            dict.Add("Mobile AMD Athlon XP", 572);
            dict.Add("Mobile AMD Athlon XP-M", 512);
            dict.Add("Mobile AMD Athlon XP-M 1400 +", 184);
            dict.Add("Mobile AMD Athlon XP-M 1500 +", 261);
            dict.Add("Mobile AMD Athlon XP-M 1600 +", 257);
            dict.Add("Mobile AMD Athlon XP-M 1700 +", 323);
            dict.Add("Mobile AMD Athlon XP-M 1800 +", 312);
            dict.Add("Mobile AMD Athlon XP-M 1900 +", 342);
            dict.Add("Mobile AMD Athlon XP-M 2000 +", 300);
            dict.Add("Mobile AMD Athlon XP-M 2200 +", 364);
            dict.Add("Mobile AMD Athlon XP-M 2400 +", 346);
            dict.Add("Mobile AMD Athlon XP-M 2500 +", 351);
            dict.Add("Mobile AMD Athlon XP-M 2600 +", 370);
            dict.Add("Mobile AMD Athlon XP-M 2800 +", 365);
            dict.Add("Mobile AMD Athlon XP-M 3000 +", 365);
            dict.Add("Mobile AMD Athlon XP-M 3100 +", 440);
            dict.Add("Mobile AMD Athlon XP-M 3200 +", 487);
            dict.Add("Mobile AMD Athlon XP-M(LV)", 508);
            dict.Add("Mobile AMD Athlon XP-M(LV) 1500 +", 288);
            dict.Add("Mobile AMD Athlon XP-M(LV) 1600 +", 296);
            dict.Add("Mobile AMD Athlon XP-M(LV) 2000 +", 335);
            dict.Add("Mobile AMD Athlon XP-M(LV) 2200 +", 373);
            dict.Add("Mobile AMD Athlon XP-M(LV) 2800 +", 475);
            dict.Add("Mobile AMD Athlon XP-M(LV) 3200 +", 505);
            dict.Add("Mobile AMD Sempron 2100 +", 252);
            dict.Add("Mobile AMD Sempron 2600 +", 394);
            dict.Add("Mobile AMD Sempron 2800 +", 358);
            dict.Add("Mobile AMD Sempron 3000 +", 388);
            dict.Add("Mobile AMD Sempron 3100 +", 416);
            dict.Add("Mobile AMD Sempron 3200 +", 354);
            dict.Add("Mobile AMD Sempron 3300 +", 444);
            dict.Add("Mobile AMD Sempron 3400 +", 443);
            dict.Add("Mobile AMD Sempron 3500 +", 411);
            dict.Add("Mobile AMD Sempron 3600 +", 428);
            dict.Add("Mobile AMD Sempron 3800 +", 517);
            dict.Add("Mobile Intel Celeron 1.20GHz", 170);
            dict.Add("Mobile Intel Celeron 1.50GHz", 188);
            dict.Add("Mobile Intel Celeron 1.60GHz", 188);
            dict.Add("Mobile Intel Celeron 1.70GHz", 181);
            dict.Add("Mobile Intel Celeron 1.80GHz", 186);
            dict.Add("Mobile Intel Celeron 2.00GHz", 204);
            dict.Add("Mobile Intel Celeron 2.20GHz", 226);
            dict.Add("Mobile Intel Celeron 2.40GHz", 239);
            dict.Add("Mobile Intel Celeron 2.50GHz", 281);
            dict.Add("Mobile Intel Celeron 1200MHz", 294);
            dict.Add("Mobile Intel Celeron 1333MHz", 309);
            dict.Add("Mobile Intel Pentium 4 2.30GHz", 278);
            dict.Add("Mobile Intel Pentium 4 2.40GHz", 264);
            dict.Add("Mobile Intel Pentium 4 2.66GHz", 279);
            dict.Add("Mobile Intel Pentium 4 2.80GHz", 315);
            dict.Add("Mobile Intel Pentium 4 3.06GHz", 337);
            dict.Add("Mobile Intel Pentium 4 3.20GHz", 356);
            dict.Add("Mobile Intel Pentium 4 3.33GHz", 439);
            dict.Add("Mobile Intel Pentium 4 3.46GHz", 478);
            dict.Add("Mobile Intel Pentium 4-M 1.60GHz", 195);
            dict.Add("Mobile Intel Pentium 4-M 1.70GHz", 173);
            dict.Add("Mobile Intel Pentium 4-M 1.80GHz", 186);
            dict.Add("Mobile Intel Pentium 4-M 1.90GHz", 232);
            dict.Add("Mobile Intel Pentium 4-M 2.00GHz", 202);
            dict.Add("Mobile Intel Pentium 4-M 2.20GHz", 209);
            dict.Add("Mobile Intel Pentium 4-M 2.40GHz", 213);
            dict.Add("Mobile Intel Pentium 4-M 2.50GHz", 256);
            dict.Add("Mobile Intel Pentium 4-M 2.60GHz", 313);
            dict.Add("Mobile Intel Pentium III-M 866MHz", 211);
            dict.Add("Mobile Intel Pentium III-M 933MHz", 200);
            dict.Add("Mobile Intel Pentium III-M 1000MHz", 250);
            dict.Add("Mobile Intel Pentium III-M 1133MHz", 289);
            dict.Add("Mobile Intel Pentium III-M 1200MHz", 263);
            dict.Add("Mobile Intel Pentium III-M 1333MHz", 289);
            dict.Add("Quad-Core AMD Opteron 1385", 3442);
            dict.Add("Quad-Core AMD Opteron 1389", 3046);
            dict.Add("VIA C3 Ezra", 100);
            dict.Add("VIA C7 1500MHz", 284);
            dict.Add("VIA C7-D 1500MHz", 287);
            dict.Add("VIA C7-D 1800MHz", 202);
            dict.Add("VIA C7-D 2000MHz", 368);
            dict.Add("VIA C7-M 1000MHz", 164);
            dict.Add("VIA C7-M 1200MHz", 141);
            dict.Add("VIA C7-M 1600MHz", 185);
            dict.Add("VIA C7-M 6300MHz", 222);
            dict.Add("VIA Eden 800MHz", 155);
            dict.Add("VIA Eden X2 U4200 @ 1.0 + GHz", 480);
            dict.Add("VIA Eden X4 C4250 @ 1.2 + GHz", 1147);
            dict.Add("VIA Esther 1000MHz", 169);
            dict.Add("VIA Esther 1300MHz", 250);
            dict.Add("VIA Esther 1500MHz", 219);
            dict.Add("VIA Esther 2000MHz", 394);
            dict.Add("VIA Nano L1900@1400MHz", 457);
            dict.Add("VIA Nano L2007@1600MHz", 381);
            dict.Add("VIA Nano L2100@1800MHz", 428);
            dict.Add("VIA Nano L2207@1600MHz", 296);
            dict.Add("VIA Nano L3025@1600MHz", 517);
            dict.Add("VIA Nano L3050@1800MHz", 581);
            dict.Add("VIA Nano L3600@2000MHz", 645);
            dict.Add("VIA Nano U2250(1.6GHz Capable)", 377);
            dict.Add("VIA Nano U2250@1300 + MHz", 274);
            dict.Add("VIA Nano U2500@1200MHz", 306);
            dict.Add("VIA Nano U3100(1.6GHz Capable)", 501);
            dict.Add("VIA Nano U3300@1200MHz", 409);
            dict.Add("VIA Nano U3500@1000MHz", 275);
            dict.Add("VIA Nano X2 U4025 @ 1.2 GHz", 639);
            dict.Add("VIA Nehemiah", 155);
            dict.Add("VIA QuadCore C4650@2.0GHz", 1940);
            dict.Add("VIA QuadCore L4700 @ 1.2 + GHz", 1363);
            dict.Add("VIA QuadCore U4650 @ 1.0 + GHz", 1173);

            int value = dict.FirstOrDefault(x => x.Key.Contains(model)).Value;
            return value;
        }

        public class CustomException : Exception
        {
            public CustomException(string msg, Exception inner) : base(msg, inner) { }
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


    }
}
