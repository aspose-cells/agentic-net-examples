// Title: Conditionally protect an Excel workbook with Aspose.Cells in C#
// Description: Loads an existing XLSX file using Aspose.Cells, checks Workbook.Settings.IsProtected, applies full password protection only if the workbook is unprotected, saves the result, and includes file‑existence validation, exception handling, and proper resource disposal.
// Keywords: Aspose.Cells | C# | .NET | protect workbook | Workbook.Settings.IsProtected | ProtectionType.All | Excel password protection | conditional workbook protection | Excel file security | avoid double protection
// Common Searches: Aspose.Cells protect workbook only if not already protected | C# check Excel workbook protection status before applying password | How to use Settings.IsProtected with Aspose.Cells | Conditional password protection for Excel files in .NET | Skip protecting already secured workbook Aspose.Cells
// Developer Intent: Add password protection to an Excel workbook only when it has no existing protection.
// Use Cases: Batch‑apply corporate password policies to a collection of Excel files while leaving already secured files untouched. | Validate incoming Excel uploads in a web service and enforce protection only for unprotected workbooks. | Create a maintenance script that audits and secures workbooks without overriding existing security settings.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, checks Settings.IsProtected, and applies ProtectionType.All with a given password only if the workbook is unprotected. | Write a robust Aspose.Cells routine that verifies the input file exists, handles exceptions, and disposes the Workbook object correctly. | Refactor the example to accept an array of file paths, protect each workbook conditionally, and output a summary report of actions performed.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing XLSX file using Aspose.Cells, checks Workbook.Settings.IsProtected, applies full password protection only if the workbook is unprotected, saves the result, and includes file‑existence validation, exception handling, and proper resource disposal.
    public class ProtectWorkbookIfNotAlreadyProtected
    {
        public static void Run()
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);

                // Check if the workbook structure or window is already protected
                if (!workbook.Settings.IsProtected)
                {
                    // Apply protection with a password
                    workbook.Protect(ProtectionType.All, "MySecretPassword");
                    Console.WriteLine("Workbook was not protected and has now been protected.");
                }
                else
                {
                    Console.WriteLine("Workbook is already protected; skipping protection.");
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorkbookIfNotAlreadyProtected.Run();
        }
    }
}
