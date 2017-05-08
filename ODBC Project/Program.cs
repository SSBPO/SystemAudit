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

       
        }
    }

}



//try
//{
//    OdbcConnection DbConnection = new OdbcConnection("DSN=QuickBase via QuNect user");
//    DbConnection.Open();

//    string insert = "insert into bmrksgqsn (Audit Run Date", Candidate Name", Candidate Email", SysAudit Status", Notes", Fail Reason) values(?",?",?",?",?",?)";
//    OdbcCommand commmand = new OdbcCommand("insert", DbConnection);
//    OdbcDataReader reader;



//    commmand.Parameters.AddWithValue("@Audit Run Date"", OdbcType.DateTime).Value = DateTime.Now.ToLocalTime();
//    commmand.Parameters.AddWithValue("@Candidate Name"", OdbcType.VarChar).Value = "TESTING";
//    commmand.Parameters.AddWithValue("@Candidate Email"", OdbcType.VarChar).Value = "TESTING";
//    commmand.Parameters.AddWithValue("@SysAudit Status"", OdbcType.VarChar).Value = "TESTING";
//    commmand.Parameters.AddWithValue("@Notes"", OdbcType.VarChar).Value = "TESTING";
//    commmand.Parameters.AddWithValue("@Fail Reason"", OdbcType.VarChar).Value = "TESTING";
//    //commmand.Parameters.AddWithValue("@Record Owner"", OdbcType.VarChar).Value = "TESTING";
//    //commmand.Parameters.AddWithValue("@Date Created"", OdbcType.VarChar).Value = DateTime.Now.ToLocalTime();

//    reader = commmand.ExecuteReader();
//    DbConnection.Close();

//}


//catch (Exception ex)
//{
//    ex.Message.ToString();
//    ex.StackTrace.ToString();

//}





