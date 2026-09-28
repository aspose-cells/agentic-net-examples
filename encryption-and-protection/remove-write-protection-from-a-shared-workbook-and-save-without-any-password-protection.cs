// Title: How to remove write protection from a shared Excel workbook and save an unprotected copy using Aspose.Cells for .NET
// AI Prompts: Load a shared .xlsx file with Aspose.Cells, call Workbook.Unprotect(string.Empty) to clear any write‑protection, and save the result to a new file path. | In C#, open a protected workbook, catch potential Unprotect exceptions, and write the workbook out without any password using Aspose.Cells.
// Common Searches: Aspose.Cells C# remove write protection from shared workbook without password | How to unprotect an Excel file programmatically with Aspose.Cells .NET | Save an unprotected copy of a protected .xlsx using Aspose.Cells | C# code to load a shared Excel workbook and disable write protection | Aspose.Cells unprotect workbook when password is unknown
// Tags: Aspose.Cells workbook unprotect | C# remove Excel write protection | save unprotected .xlsx Aspose.Cells | shared workbook protection .NET | Workbook.Unprotect empty password

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example verifies the input file, loads the shared workbook with Aspose.Cells, attempts to clear write protection by calling Unprotect with an empty string (safely handling any exception if the workbook is already unprotected), and then saves the workbook to a new location without any password protection.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "SharedWorkbook.xlsx";
            const string outputPath = "SharedWorkbook_Unprotected.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing shared workbook
                var workbook = new Workbook(inputPath);

                // Attempt to remove write‑protection (empty password works when no password is set)
                try
                {
                    workbook.Unprotect(string.Empty);
                }
                catch (Exception ex)
                {
                    // If the workbook is not protected, Unprotect may throw; ignore safely
                    Console.WriteLine($"Unprotect warning: {ex.Message}");
                }

                // Save the workbook without any protection
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
