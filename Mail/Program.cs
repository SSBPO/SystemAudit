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
using System.Net.NetworkInformation;
using System.Threading;
using System.Net;

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

            DateTime date = DateTime.Now;
            Console.WriteLine(date.ToString("htt"));


        }



    }
}
