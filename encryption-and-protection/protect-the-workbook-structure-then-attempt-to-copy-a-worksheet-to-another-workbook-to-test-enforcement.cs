using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook sourceWorkbook = new Workbook();                     // create new workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];        // get first worksheet
            sourceSheet.Name = "DataSheet";
            sourceSheet.Cells["A1"].PutValue("Protected Sheet");         // add sample data

            // Protect the workbook structure with a password
            sourceWorkbook.Protect(ProtectionType.Structure, "pwd123");

            // Save the protected workbook (optional, just for verification)
            sourceWorkbook.Save("ProtectedSource.xlsx", SaveFormat.Xlsx);

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();                       // create empty workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];            // get first (default) sheet
            destSheet.Name = "CopyTarget";

            // Attempt to copy the protected worksheet into the destination workbook
            try
            {
                // This operation should respect the protection and may throw an exception
                destSheet.Copy(sourceSheet);
                Console.WriteLine("Worksheet copied successfully despite protection.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to copy worksheet due to protection:");
                Console.WriteLine(ex.Message);
            }

            // Save the destination workbook to observe the result
            destWorkbook.Save("CopyAttemptResult.xlsx", SaveFormat.Xlsx);
        }
    }
}