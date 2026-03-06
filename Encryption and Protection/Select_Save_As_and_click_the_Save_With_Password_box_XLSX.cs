using System;
using Aspose.Cells;

namespace AsposeCellsPasswordSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set a password to protect the workbook when saving
            workbook.Settings.Password = "MySecurePassword";

            // Save the workbook as XLSX with the password protection applied
            workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: Verify that the workbook is password protected
            Console.WriteLine("Workbook saved with password protection: " + workbook.IsWorkbookProtectedWithPassword);
        }
    }
}