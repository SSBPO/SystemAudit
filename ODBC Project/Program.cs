using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;
using System.Data;

namespace ODBC_Project
{
    class Program
    {

        static void Main(string[] args)
        {
            try
            {
                OdbcConnection DbConnection = new OdbcConnection("DSN=QuickBase via QuNect user");
                DbConnection.Open();

                string insert = "insert into bmrksgqsn (Audit Run Date, Candidate Name, Candidate Email, SysAudit Status, Notes, Fail Reason) values(?,?,?,?,?,?)";
                OdbcCommand commmand = new OdbcCommand(insert, DbConnection);
                OdbcDataReader reader;

                


                commmand.Parameters.AddWithValue("@Audit Run Date", OdbcType.DateTime).Value = DateTime.Now.ToLocalTime();

                commmand.Parameters.AddWithValue("@Candidate Name", OdbcType.VarChar).Value = "Test Testing";

                commmand.Parameters.AddWithValue("@Candidate Email", OdbcType.VarChar).Value = "Test@test.com";

                commmand.Parameters.AddWithValue("@SysAudit Status", OdbcType.VarChar).Value = "Fail";

                commmand.Parameters.AddWithValue("@Notes", OdbcType.VarChar).Value = "[OS = Windows 10 Home Build 14393] [CPU Score = 7] - [Processor = AMD A8-6410 APU with AMD Radeon R5 Graphics ] - [RAM Score = 5.9] - [Total RAM = 3GB] - [Disk Score = 5.9] - [Total Space = 675GB] - [Available Space = 589GB] - [Download Speed = 10 Mbps] - [Upload Speed = 3.43 Mbps]";
                                                                                    
                commmand.Parameters.AddWithValue("@Fail Reason", OdbcType.VarChar).Value = "CPU Insufficient";


                reader = commmand.ExecuteReader();
                DbConnection.Close();
                Console.WriteLine("Completed import");

            }


            catch (Exception ex)
            {
                ex.Message.ToString();
                ex.StackTrace.ToString();

            }
        }
    }

}





//}





