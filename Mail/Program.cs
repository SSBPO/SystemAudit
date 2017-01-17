using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
//using AE.Net.Mail;
using System.Web;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using Microsoft.Vbe.Interop;
using System.Diagnostics;

namespace Mail
{
    class Program
    {
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;

        static void Main(string[] args)
        {
            //Outlook.Account OAccount = new Outlook.Account();
            //foreach(OAccount in Outlook.Session.Accounts)
            //{

            //}



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

        public static void writeTogExcel(string dateProccessed, string Name, string Email, string reseults, string results, string processedBy)
        {
            string path = @"C:\Winaudit\All SystemAudits Results.csv";

            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;

            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);          
            mWorkSheets = mWorkBook.Worksheets;
            mWSheet1 = (Excel.Worksheet)mWorkSheets.get_Item("All SystemAudits Results");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;
            int index = rowCount + 1;

            mWSheet1.Cells[index, 1] = dateProccessed;
            mWSheet1.Cells[index, 2] = Name;
            mWSheet1.Cells[index, 3] = Email;
            mWSheet1.Cells[index, 4] = reseults;
            mWSheet1.Cells[index, 5] = processedBy;

            mWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }


        private static void KillExcel()
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
    }
}
