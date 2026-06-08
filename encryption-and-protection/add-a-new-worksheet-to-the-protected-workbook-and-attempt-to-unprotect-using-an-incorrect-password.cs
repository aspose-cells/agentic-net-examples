using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Protect the workbook structure with a password
                workbook.Protect(ProtectionType.Structure, "correctPassword");

                // Add a new worksheet (allowed programmatically even when protected)
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                newSheet.Name = "NewSheet";

                // Attempt to unprotect with an incorrect password
                try
                {
                    workbook.Unprotect("wrongPassword");
                    Console.WriteLine("Workbook unprotected with wrong password (unexpected).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to unprotect workbook with wrong password: " + ex.Message);
                }

                // Save the workbook to verify the state (still protected)
                string outputPath = "ProtectedWorkbookDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}