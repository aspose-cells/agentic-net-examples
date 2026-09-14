// Title: Apply password protection to workbook structure with Aspose.Cells for .NET and verify that new worksheets cannot be added
// AI Prompts: Generate C# code that applies password‑based structure protection to an Aspose.Cells workbook. | Write a try‑catch block that attempts to add a new worksheet after the workbook is structure‑protected and logs the error. | Show how to save the password‑protected workbook to an .xlsx file and confirm that the protection is active.
// Common Searches: Aspose.Cells C# example for protecting only the workbook structure with a password | How to prevent adding new sheets after applying structure protection in Aspose.Cells | Exception thrown when inserting worksheet into a structure‑protected Excel file using Aspose.Cells | Saving a password‑protected Excel workbook with Aspose.Cells .NET
// Tags: Workbook structure protection password Aspose.Cells | Block worksheet insertion after protection C# | Aspose.Cells Protect method structure only | Handle protection exception when adding worksheet | Save password‑protected Excel file .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, adds two worksheets, protects only the workbook structure with a password via workbook.Protect(ProtectionType.Structure, password), attempts to add a third worksheet which triggers an exception, catches and logs the error, and finally saves the password‑protected workbook as 'ProtectedWorkbook.xlsx'.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample worksheets
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");

                // Protect the workbook structure with a password (structure only)
                string password = "MySecretPassword";
                workbook.Protect(ProtectionType.Structure, password);

                // Attempt to add a new worksheet to demonstrate protection enforcement
                try
                {
                    workbook.Worksheets.Add("NewSheet");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Unable to add worksheet: " + ex.Message);
                }

                // Save the workbook (optional)
                string outputPath = "ProtectedWorkbook.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error saving workbook: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }
}
