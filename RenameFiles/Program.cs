using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Data;
using System.Linq;

namespace RenameFiles
{
    class Program
    {
        static void Main(string[] args)
        {

            RenameFiles(@"C:\Downloads", "WTC");
        }

        static void RenameFiles(string dirName, string cusName)
        {

            var dirnames = Directory.GetDirectories(dirName);

            int i = 0;

            try
            {
                foreach (var dir in dirnames)
                {
                    var fnames = Directory.GetFiles(dir, "*.mp3").Select(Path.GetFileName);

                    DirectoryInfo d = new DirectoryInfo(dir);
                    FileInfo[] finfo = d.GetFiles("*.mp3");

                    foreach (var f in fnames)
                    {
                        i++;
                        Console.WriteLine(d + @"\" + f.ToString());

                        if (!File.Exists(f.ToString()))
                        {
                            File.Move(d + @"\" + f.ToString(), d + @"\" + cusName + " - " + f.ToString());
                        }
                        else
                        {
                            Console.WriteLine("The file you are attempting to rename already exists! The file path is {0}.", dir);
                            foreach (FileInfo fi in finfo)
                            {
                                Console.WriteLine("The file modify date is: {0} ", File.GetLastWriteTime(dir));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //  Console.Read();

        }
    }
}
