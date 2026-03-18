using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UpdateSqlDataConnectionDemo
    {
        public static void Run()
        {
            // Create a new workbook (default XLSX)
            Workbook workbook = new Workbook();

            // (Optional) Add any data or worksheets here as needed.

            // Save the workbook
            workbook.Save("UpdatedSqlDataConnection.xlsx");

            Console.WriteLine("Workbook saved successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateSqlDataConnectionDemo.Run();
        }
    }
}