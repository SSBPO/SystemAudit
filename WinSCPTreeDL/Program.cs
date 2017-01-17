using System;
using System.Collections.Generic;
using System.IO;
using WinSCP;

using GemBox.Spreadsheet;
using System.Text;
using System.Data;
using System.Linq;

namespace WinSCPTreeDL
{

    class FolderDLApp
    {
        static void Main()
        {

            string excelLicense = "EQU2-1K6F-UZPP-4MOX";
            string SFTPTransferResult = "";
            int configFiles = 0;

            ConfigData xxx_configData = new ConfigData();
            FolderDLApp xxx_downLoadApp = new FolderDLApp();

            SpreadsheetInfo.SetLicense(excelLicense);    
            Console.Clear();

            Console.WriteLine("############################################################################################");
            Console.WriteLine("################################# FolderDLApp ver 1.0 ######################################");
            Console.WriteLine("############################################################################################");
            Console.WriteLine();
            Console.WriteLine();
           
            string[] fileEntries = Directory.GetFiles(xxx_configData.configFDpath);

            if(fileEntries.Count() > 1)
                Console.WriteLine("There are " + fileEntries.Count() + " configuration files in the directory.");
            else
            {
                Console.WriteLine("There is " + fileEntries.Count() + " configuration file in the directory.");
            }

            foreach (string file in fileEntries)
            {
                configFiles = configFiles + 1;
                xxx_configData = FolderDLApp.readConfig(file);
               
                Console.WriteLine();
                Console.WriteLine("Processing file " + configFiles);
                Console.WriteLine();
                Console.WriteLine("     SFTPTransType:  " + xxx_configData.SFTPTransType.Replace("1", "Download"));
                Console.WriteLine("     SFTPHostName:   " + xxx_configData.SFTPHostName);
                Console.WriteLine("     SFTPUserName:   " + xxx_configData.SFTPUserName);
                Console.WriteLine("     SFTPPassword:   " + xxx_configData.SFTPPassword);
                Console.WriteLine("     SFTPCusName:    " + xxx_configData.SFTPCusName);
                Console.WriteLine("     SFTPlocalPath:  " + xxx_configData.SFTPlocalPath);
                Console.WriteLine("     SFTPremotePath: " + xxx_configData.SFTPremotePath);
                Console.WriteLine();

                switch (xxx_configData.SFTPTransType)
                {
                    case "1":
                        Console.WriteLine("Downloading files from TFTP server...");
                        Console.WriteLine();
                        SFTPTransferResult = SFTPTransfers.SFTPDownload(xxx_configData.SFTPHostName, xxx_configData.SFTPUserName, xxx_configData.SFTPCusName, xxx_configData.SFTPPassword, xxx_configData.SFTPlocalPath, xxx_configData.SFTPremotePath, xxx_configData.SFTPSshHostKey);
                        Console.WriteLine();
                        break;
                    case "2":
                        Console.WriteLine("Uploading files to TFTP server...");
                        Console.WriteLine();
                        SFTPTransferResult = SFTPTransfers.SFTPUploads(xxx_configData.SFTPHostName, xxx_configData.SFTPUserName, xxx_configData.SFTPCusName, xxx_configData.SFTPPassword, xxx_configData.SFTPlocalPath, xxx_configData.SFTPremotePath, xxx_configData.SFTPSshHostKey);
                        Console.WriteLine();
                        break;
                    case "3":
                        Console.WriteLine("Sychrinizing local folders to remote TFTP server...");
                        Console.WriteLine();
                        SFTPTransferResult = SFTPTransfers.SFTPsynchronize(xxx_configData.SFTPHostName, xxx_configData.SFTPUserName, xxx_configData.SFTPCusName, xxx_configData.SFTPPassword, xxx_configData.SFTPlocalPath, xxx_configData.SFTPremotePath, xxx_configData.SFTPSshHostKey);
                        Console.WriteLine();
                        break;
                    default:
                        Console.WriteLine("Sorry, invalid transaction type.");
                        break;
                }

                SFTPTransfers.RenameFiles(xxx_configData.SFTPlocalPath, xxx_configData.SFTPCusName);
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine(SFTPTransferResult);
                Console.WriteLine("Transaction completed.");
                Console.WriteLine();
                
            }

            
            Console.ReadLine();

        }

        private static void FileTransferred(object sender, TransferEventArgs e)
        {
            if (e.Error == null)
            {
                Console.WriteLine("Upload of {0} succeeded", e.FileName);
            }
            else
            {
                Console.WriteLine("Upload of {0} failed: {1}", e.FileName, e.Error);
            }

            if (e.Chmod != null)
            {
                if (e.Chmod.Error == null)
                {
                    Console.WriteLine("Permisions of {0} set to {1}", e.Chmod.FileName, e.Chmod.FilePermissions);
                }
                else
                {
                    Console.WriteLine("Setting permissions of {0} failed: {1}", e.Chmod.FileName, e.Chmod.Error);
                }
            }
            else
            {
                Console.WriteLine("Permissions of {0} kept with their defaults", e.Destination);
            }

            if (e.Touch != null)
            {
                if (e.Touch.Error == null)
                {
                    Console.WriteLine("Timestamp of {0} set to {1}", e.Touch.FileName, e.Touch.LastWriteTime);
                }
                else
                {
                    Console.WriteLine("Setting timestamp of {0} failed: {1}", e.Touch.FileName, e.Touch.Error);
                }
            }
            else
            {
                // This should never happen during "local to remote" synchronization
                Console.WriteLine("Timestamp of {0} kept with its default (current time)", e.Destination);
            }
        }

        public static ConfigData readConfig(string configFilePath)
        {
            //string[] fileEntry = Directory.GetFiles(configFilePath);
            ConfigData xxx_configData = new ConfigData();

                DataTable vendorData = XLSXImport.GetDataFromXLSX(configFilePath);

                foreach (DataRow dataItem in vendorData.Rows)
                {
                    string dataOption = (dataItem["Option"].ToString().Trim());
                    string dataValue = (dataItem["Value"].ToString().Trim());
                

                    switch (dataOption)
                    {
                        case "SFTPTransType":
                            xxx_configData.SFTPTransType = dataValue;
                            break;
                        case "ConfigFolder":
                        xxx_configData.configFDpath = dataValue;
                        break;
                        case "SFTPHostName":
                            xxx_configData.SFTPHostName = dataValue;
                            break;
                        case "SFTPSshHostKey":
                            xxx_configData.SFTPSshHostKey = dataValue;
                            break;
                        case "SFTPUserName":
                            xxx_configData.SFTPUserName = dataValue;
                            break;
                        case "SFTPPassword":
                            xxx_configData.SFTPPassword = dataValue;
                            break;
                        case "SFTPCusName":
                            xxx_configData.SFTPCusName = dataValue;
                            break;
                        case "SFTPlocalPath":
                                // Create local subdirectory, if it does not exist yet
                                if (!Directory.Exists(dataValue))
                                {
                            Directory.Delete(dataValue);
                            Directory.CreateDirectory(dataValue);
                                }
                        xxx_configData.SFTPlocalPath = dataValue;
                            break;
                        case "SFTPremotePath":
                            xxx_configData.SFTPremotePath = dataValue;
                            break;
                        default:
                            break;
                    }

            }

            return xxx_configData;
        }

        public class ConfigData
        {
            public string configFDpath = @"C:\config";            

            public string SFTPTransType { get; set; }
            public string SFTPHostName { get; set; }
            public string SFTPCusName { get; set; }
            public string SFTPUserName { get; set; }
            public string SFTPPassword { get; set; }
            public string SFTPlocalPath { get; set; }
            public string SFTPremotePath { get; set; }
            public string SFTPSshHostKey { get; set; }       
        }

        public class XLSXImport
        {
            public static DataTable GetDataFromXLSX(string fileEntry)
            {
                // Load Excel file.
                ExcelFile ef = ExcelFile.Load(fileEntry);

                // Select the first worksheet from the file.
                ExcelWorksheet ws = ef.Worksheets[0];

                DataTable XLSXData = ws.CreateDataTable(new CreateDataTableOptions()
                {
                    ColumnHeaders = true,
                    StartRow = 0,
                    NumberOfColumns = 8,
                    NumberOfRows = ws.Rows.Count,
                    Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
                });

                // Write DataTable content
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("DataTable content:");
                foreach (DataRow row in XLSXData.Rows)
                {
                    sb.AppendFormat("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", row[0], row[1], row[2], row[3], row[4], row[5]);
                    sb.AppendLine();
                }

                //  Console.WriteLine(sb.ToString());
                return XLSXData;
            }

        }

        public class SFTPTransfers
        {

            public static string SFTPDownload(string SFTPHostName, string SFTPUserName, string SFTPCustName, string SFTPPassword, string SFTPlocalPath, string SFTPremotePath, string SFTPHostKey)
            {
                try
                {

                    string newDLFolder = string.Empty;

                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Sftp,
                        HostName = SFTPHostName,
                        UserName = SFTPUserName,
                        Password = SFTPPassword,
                        SshHostKeyFingerprint = SFTPHostKey
                    };

                    using (Session session = new Session())
                    {
                        session.FileTransferred += FileTransferred;
                        session.Open(sessionOptions);

                        // Enumerate files and directories to download
                        IEnumerable<RemoteFileInfo> fileInfos =
                            session.EnumerateRemoteFiles(
                                SFTPremotePath, null,
                                EnumerationOptions.EnumerateDirectories | EnumerationOptions.AllDirectories);

                        foreach (RemoteFileInfo fileInfo in fileInfos)
                        {
                            string localFilePath = session.TranslateRemotePathToLocal(fileInfo.FullName, SFTPremotePath, SFTPlocalPath);

                            if (fileInfo.IsDirectory)
                            {
                                // Create local subdirectory, if it does not exist yet
                                if (!Directory.Exists(localFilePath))
                                {
                                    Directory.CreateDirectory(localFilePath);
                                    newDLFolder = localFilePath;

                                }
                            }

                            // A static method is also available.
                            bool areEqual = String.Equals(localFilePath, newDLFolder, StringComparison.Ordinal);

                            if (areEqual)
                            {
                                Console.WriteLine(string.Format("Downloading file {0}...", fileInfo.Name));
                                Console.WriteLine();

                                // Download file
                                var transferResult =
                                    session.GetFiles(session.EscapeFileMask(fileInfo.FullName), localFilePath);

                                // Did the download succeeded?
                                if (!transferResult.IsSuccess)
                                {
                                    // Print error (but continue with other files)
                                    Console.WriteLine(string.Format("Error downloading file {0}: {1}", fileInfo.FullName, transferResult.Failures[0].Message));
                                }
                            }
                        }

                        Console.WriteLine("Download completed. ");
                        Console.WriteLine("Renaming files...");

                        RenameFiles(SFTPlocalPath, SFTPCustName);
                        Console.WriteLine("Renaming files was sucessful!");
                    }


                    return "Download successful.";
                }

                catch (Exception e)
                {
                    Console.WriteLine("Error: {0}", e);
                    return "Download Unsuccessful";
                }


            }

            public static string SFTPUploads(string SFTPHostName, string SFTPUserName, string SFTPCustName, string SFTPPassword, string SFTPlocalPath, string SFTPremotePath, string SFTPHostKey)
            {
                try
                {

                    string newDLFolder = string.Empty;

                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Sftp,
                        HostName = SFTPHostName,
                        UserName = SFTPUserName,
                        Password = SFTPPassword,
                        SshHostKeyFingerprint = SFTPHostKey
                    };


                    using (Session session = new Session())
                    {
                        session.FileTransferred += FileTransferred;
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;
                        TransferOperationResult transferResult;
                        transferResult = session.PutFiles(SFTPlocalPath, SFTPremotePath, false, transferOptions);
                        transferResult.Check();
                        //Add try catch
                        foreach (TransferEventArgs transfer in transferResult.Transfers)
                        {
                            Console.WriteLine("Upload of {0} succeeded", transfer.Touch);
                            Console.WriteLine();
                        }

                    }

                    Console.WriteLine("Upload complete!");
                    //  Console.ReadLine();

                    return "Upload successful.";
                }

                catch (Exception e)
                {
                    Console.WriteLine("Error: {0}", e);
                    return "Upload Unsuccessful";
                }


            }

            public static string SFTPsynchronize(string SFTPHostName, string SFTPUserName, string SFTPCustName, string SFTPPassword, string SFTPlocalPath, string SFTPremotePath, string SFTPHostKey)
            {
                try
                {

                    string newDLFolder = string.Empty;

                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Sftp,
                        HostName = SFTPHostName,
                        UserName = SFTPUserName,
                        Password = SFTPPassword,
                        SshHostKeyFingerprint = SFTPHostKey
                    };

                    Console.WriteLine("Synchronizing files...");

                    using (Session session = new Session())
                    {
                        session.FileTransferred += FileTransferred;
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;
                        SynchronizationResult syncResult;
                        syncResult = session.SynchronizeDirectories(SynchronizationMode.Both, SFTPlocalPath, SFTPremotePath, false);

                        foreach (TransferEventArgs transfer in syncResult.Downloads)
                        {
                            Console.WriteLine();
                            Console.WriteLine("{0} folders were synchronized.", transfer.FileName.Count());
                            // System.Threading.Thread.Sleep(5000);
                            syncResult.Check();
                        }
                    }

                    Console.WriteLine("Sychronizing complete!");
                    RenameFiles(SFTPlocalPath, SFTPCustName);

                    return "Sychronizing was successful";

                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: {0}", e);
                    return "Synchronization Unsuccessful";
                }

            }
            
            public static string RenameFiles(string dirName, string cusName)
            {

                var dirNames = Directory.GetDirectories(dirName);

                try
                {
                    foreach (var dir in dirNames)
                    {
                       // var fnames = Directory.GetFiles(dir, "*.mp3").Select(Path.GetFileName);
                        var fnames = Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".mp3") || s.EndsWith(".wav") || s.EndsWith(".wma") || s.EndsWith(".ra") || s.EndsWith(".m4p"));                 

                        DirectoryInfo d = new DirectoryInfo(dir);                        
                        FileInfo[] finfo = d.GetFiles();


                        foreach (var f in fnames)
                        {
                            Console.WriteLine(d + @"\" + f.ToString());

                            if (!File.Exists(f.ToString()))
                            {
                                if (!f.ToString().Contains(cusName));                                
                                    File.Move(d + @"\" + f.ToString(), d + @"\" + cusName + " - " + f.ToString());
                               
                            }
                            else
                            {
                                Console.WriteLine("File not found.", dir + @"\" + f.ToString());

                                foreach (FileInfo fi in finfo)
                                {
                                    Console.WriteLine("Last modify date: {0} ", File.GetLastAccessTime(dir));
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: {0}", e);
                    return "Renaming of files was unsuccessful";
                }
                
                return "Files were renamed successfully!.";
            }

        }

     }
}




