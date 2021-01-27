using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using QiHe.CodeLib;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ganss.Excel;
using System.Data.SqlClient;









//Author: Armin Prutina
//Date: 1/26/2021

namespace Introduction_ConsoleApp_Excel
{
    class Excel1
    {
        private static void Main(string[] args)
        {


            //Creating the connection string for your SQL Server delete the next comment and insert your own

            //string connectionString = "Server=(YOUR SERVER NAME)\\SQLEXPRESS;Database=(YOUR DATABASE);User ID=(YOUR USERNAME);(YOUR PASSWORD);


            //opens up the connection
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {//selects from your table
                    cmd.Connection = conn;
                    cmd.CommandText = "SELECT * FROM Contacts";

                    SqlDataReader dr = cmd.ExecuteReader();

                    while (dr.Read())
                    {//converts the data from the SQL file to string
                        string firstName = dr["FirstName"].ToString();
                        string lastName = dr["LastName"].ToString();
                        //reads the data from SQL file back to you
                        Console.WriteLine(firstName + " " + lastName);

                    }//closes the databaseconnection
                    dr.Close();
                }
            }
            Console.ReadKey();
        }

    }
}
