using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WriteProtectionDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Set the password required to modify the workbook (enables write protection)
            workbook.Settings.WriteProtection.Password = "modifyPwd";

            // Save the workbook as an XLSX file
            workbook.Save("WriteProtectedWorkbook.xlsx");
        }
    }
}