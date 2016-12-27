using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using AE.Net.Mail;
using AE.Net.Mail.Imap;
using System.Configuration;
using Microsoft.Office.Core;
using System.Web;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Reflection;
using System.Diagnostics;

namespace WinA
{
    class Program
    {


        static void Main(string[] args)
        {
            Console.WriteLine("############################################################");
            Console.WriteLine("################# WinAuditPro ver 5 ########################");
            Console.WriteLine("############################################################");
            Console.WriteLine("");

            Program SSWinAudit = new Program();
            HtmlToText stripHtml = new HtmlToText();
            Excell excelWrite = new Excell();
            KillExcel();


            int NoEmailsProcessed = 0;


            Console.WriteLine("#### Connecting to WinAudit mailbox.........................");
            Console.WriteLine("............................................................");

            using (var imapClient = new ImapClient("secure.emailsrvr.com", "winaudit@email.jlodge.com", "W31is+en2016", AuthMethods.Login, 993, true))
            {
                imapClient.SelectMailbox("INBOX");
                Regex regex = new Regex(@":");

                var mS = imapClient.SearchMessages(SearchCondition.Unseen(), true);

                foreach (var m in mS.ToList())
                {
                    KillExcel();
                    NoEmailsProcessed = NoEmailsProcessed + 1;

                    int index = 1;

                    string cEmail = "";
                    string cName = "";
                    string cOS = "";
                    string cCPUScore = "";
                    string cRAM = "";
                    string cInternetUp = "";
                    string cInternetDown = "";
                    string cHDD = "";


                    Excel.Application oXL = new Excel.Application();
                    Excel.Workbook oWB = oXL.Workbooks.Open(@"C:\Winaudit\WinAuditPro.xltm");
                    Excel.Worksheet oWS = oWB.Worksheets[1] as Excel.Worksheet;

                    // oXL.Visible = true;

                    Console.WriteLine("#### Collecting WinAudit data from mailbox............");
                    Console.WriteLine("......................................................");


                    try
                    {
                        string[] line = stripHtml.Convert(m.ToString()).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                        foreach (string l in line.ToList())
                        {

                            if (l != "")
                            {
                                oXL.Cells[index, 2] = l.ToString().Trim().Replace("Simplified Audit Results", "");
                                index++;

                                if (l.Contains("Name"))
                                {
                                    cName = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("Email:"))
                                {
                                    cEmail = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("Hard Drive"))
                                {
                                    cHDD = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("OS"))
                                {
                                    cOS = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("RAM"))
                                {
                                    cRAM = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("Upload"))
                                {
                                    cInternetUp = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("Download"))
                                {
                                    cInternetDown = l.ToString().Substring(0, 17);
                                }

                                if (l.Contains("CPU"))
                                {
                                    cCPUScore = l.ToString().Substring(0, 17);
                                }

                            }

                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                    oXL.Run("SaveAsC");
                    oXL.Run("SaveAsR");

                    Console.WriteLine("Reading " + cName);
                    Console.WriteLine("#### Exporting Candidate " + cName + " worksheet as .PDF... ");

                }


                imapClient.Disconnect();

            }

            //  SendMail(string recipientt, string subject, string body, string attachmentFilename);

            Console.WriteLine("");
            Console.WriteLine("#### Exporting to .PDF was successful!");
            Console.WriteLine(NoEmailsProcessed + " WinAudit were processed.");

            Console.Read();

        }



        public static void SendMail(string recipient, string subject, string body, string attachmentFilename)
        {
            //SmtpClient smtpClient = new SmtpClient();
            //NetworkCredential basicCredential = new NetworkCredential(MailConst.Username, MailConst.Password);
            //MailMessage message = new MailMessage();
            //MailAddress fromAddress = new MailAddress(MailConst.Username);

            //// setup up the host, increase the timeout to 5 minutes
            //smtpClient.Host = MailConst.SmtpServer;
            //smtpClient.UseDefaultCredentials = false;
            //smtpClient.Credentials = basicCredential;
            //smtpClient.Timeout = (60 * 5 * 1000);

            //message.From = fromAddress;
            //message.Subject = subject;
            //message.IsBodyHtml = false;
            //message.Body = body;
            //message.To.Add(recipient);

            //if (attachmentFilename != null)
            //    message.Attachments.Add(new Attachment(attachmentFilename));

            //smtpClient.Send(message);
        }


        private void CreateItemFromTemplate()
        {
            Console.WriteLine("true");
            //open temple, populate template,attach pdf and send two messages
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

                /// <summary>
                /// Normally, extra whitespace characters are discarded.
                /// If this property is set to true, they are passed
                /// through unchanged.
                /// </summary>
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

                /// <summary>
                /// Clears all current text.
                /// </summary>
                public void Clear()
                {
                    _text.Length = 0;
                    _currLine.Length = 0;
                    _emptyLines = 0;
                }

                /// <summary>
                /// Writes the given string to the output buffer.
                /// </summary>
                /// <param name="s"></param>
                public void Write(string s)
                {
                    foreach (char c in s)
                        Write(c);
                }

                /// <summary>
                /// Writes the given character to the output buffer.
                /// </summary>
                /// <param name="c">Character to write</param>
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

                // Appends the current line to output buffer
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

                /// <summary>
                /// Returns the current output as a string.
                /// </summary>
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


        //public static List<MailMessage> ReadMail()
        //{
        //    List<MailMessage> messages = null;
        //    try
        //    {
        //        string userName = "winaudit@statesidebpo.com"; // Replace with your actual gmail id
        //        string passWord = "W31is+en2016"; // Replace with your password

        //        if (!string.IsNullOrEmpty(userName) && !string.IsNullOrEmpty(passWord))
        //        {
        //            using (var imapClient =
        //                new ImapClient("secure.emailsrvr.com", userName, passWord, AuthMethods.Login, 993, true))
        //            {
        //                imapClient.SelectMailbox("INBOX");
        //                messages = new List<MailMessage>(imapClient.GetMessageCount());
        //                messages = imapClient.GetMessages(0, 100, false, true).ToList();
        //                imapClient.Disconnect();
        //            }
        //        }
        //        else
        //        {
        //            Console.WriteLine("Username or Password is empty!");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }

        //    return messages;
        //}
    }
}
