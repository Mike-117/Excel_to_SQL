
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelApp = Microsoft.Office.Interop.Excel;//Install NuGet Package for solution
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToSqlConsole
{
    class Program
    {
        static void Main(string[] args)
        {

            //Create COM Objects.
            ExcelApp.Application excelApp = new ExcelApp.Application();
            DataRow myNewRow;
            DataTable myTable;

            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }
            
            Console.WriteLine("Please enter the full path of the excel worksheet you would like to store in a database.");//The file SpaghettiWesterns.xlsx is in the project folder.
            string excelFile = Console.ReadLine();
            ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(@excelFile);
            ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
            ExcelApp.Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            //Set DataTable Name and Columns Name
            myTable = new DataTable("SpaghettiWesterns");
            myTable.Columns.Add("Id", typeof(int));
            myTable.Columns.Add("Prefix", typeof(string));
            myTable.Columns.Add("Title", typeof(string));
            myTable.Columns.Add("Director", typeof(string));
            myTable.Columns.Add("Lead", typeof(string));
            myTable.Columns.Add("Year", typeof(int));



            //first row using for heading, start second row for data
            for (int i = 2; i <= rows; i++)
            {
                myNewRow = myTable.NewRow();
                myNewRow["Id"] = (i - 1).ToString();
                myNewRow["Prefix"] = Convert.ToString(excelRange.Cells[i, 1].Value2); //convert to string to accept null values
                myNewRow["Title"] = excelRange.Cells[i, 2].Value2.ToString(); //string
                myNewRow["Director"] = excelRange.Cells[i, 3].Value2.ToString(); //string
                myNewRow["Lead"] = excelRange.Cells[i, 4].Value2.ToString(); //string
                myNewRow["Year"] = Convert.ToInt32(excelRange.Cells[i, 5].Value2.ToString()); //integer

                myTable.Rows.Add(myNewRow);
            }
            //Display list in console
            foreach (DataRow dr in myTable.Rows)
            {
                Console.WriteLine("Title : {0} {1}\tDirector : {2}\tLead : {3}\tYear : {4}", dr[1], dr[2], dr[3], dr[4], dr[5]);
            }


            using (SqlConnection sqlconnection = new SqlConnection(@"Data Source=DESKTOP-S3B4I2C\SQLEXPRESS;Initial Catalog=Movies;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"))
            {
                sqlconnection.Open();

                // Create a Table on SQL Server 
                string createTableQuery = @"Create Table [SpaghettiWesterns] 
                        (Id int, Prefix varchar(10), Title varchar(50), Director varchar(50), Lead varchar(50), Year int)";

                SqlCommand command = new SqlCommand(createTableQuery, sqlconnection);
                command.ExecuteNonQuery();

                // Copy the DataTable to SQL Server Table using SqlBulkCopy
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlconnection))
                {
                    sqlBulkCopy.DestinationTableName = myTable.TableName;

                    foreach (var column in myTable.Columns)
                        sqlBulkCopy.ColumnMappings.Add(column.ToString(), column.ToString());

                    sqlBulkCopy.WriteToServer(myTable);
                    sqlconnection.Close();
                }



            }
            //after reading, relaase the excel project
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            
            Console.WriteLine("\n\nYou have successfully imported your excel file to a database.");
            Console.ReadLine();
        }

    }

}