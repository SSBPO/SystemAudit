using System;
using System.Collections.Generic;
using System.IO;
using WinSCP;

using GemBox.Spreadsheet;
using System.Text;
using System.Data;

namespace WinSCPTreeDL
{

    class FolderDLApp
    {
        public static object Items { get; private set; }

        static void Main()
        {
            // If using Professional version, put your serial key below.

            SpreadsheetInfo.SetLicense("EQU2-1K6F-UZPP-4MOX");

            try
            {
                // Initialize Import Variables
                ConfigData xxx_configData = new ConfigData();

                // Get CSV location from text file

                using (var csvLocationReader = new StreamReader("config.txt"))
                {
                    //
                    string line = string.Empty;

                    while (line != null)
                    {

                        line = csvLocationReader.ReadLine();
                        if (line != null)
                            xxx_configData.configFolder = line;
                        Console.WriteLine(line);
                        // Console.WriteLine(xxx_configData.configFolder);
                    }
                }

            }
            catch (FileNotFoundException e)
            {
                // FileNotFoundExceptions are handled here.  
            }
        }
    }
}

        
    
