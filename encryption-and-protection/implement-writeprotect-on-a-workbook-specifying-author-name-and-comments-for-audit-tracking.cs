using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample data
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Add an audit comment to the cell
            int commentIndex = sheet.Comments.Add("A1");
            sheet.Comments[commentIndex].Note = "Data entered by John Doe on " + DateTime.Now.ToString("yyyy-MM-dd");
            sheet.Comments[commentIndex].Author = "AuditSystem";

            // Configure write protection (author and password)
            workbook.Settings.WriteProtection.Password = "StrongPassword!123";
            workbook.Settings.WriteProtection.Author = "John Doe";
            workbook.Settings.WriteProtection.RecommendReadOnly = true; // optional read‑only recommendation

            // Save the workbook (lifecycle save rule)
            string outputPath = "WriteProtectedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Load the saved workbook to verify protection (lifecycle load rule)
            Workbook loadedWorkbook = new Workbook(outputPath);
            Console.WriteLine("Write protection author: " + loadedWorkbook.Settings.WriteProtection.Author);
            Console.WriteLine("Is write protected: " + loadedWorkbook.Settings.WriteProtection.IsWriteProtected);
        }
    }
}