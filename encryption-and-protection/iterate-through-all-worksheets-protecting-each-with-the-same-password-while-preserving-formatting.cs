// Title: Protect all worksheets with one password using Aspose.Cells for .NET
// Description: Loads an existing workbook, loops through every Worksheet, applies ProtectionType.All with a shared password, and saves the file while keeping all cell, row, column, and style formatting intact. Includes file‑existence check and exception handling.
// Keywords: Aspose.Cells worksheet protection | C# protect all sheets | Excel password protection .NET | preserve formatting Aspose | bulk sheet protect Aspose.Cells | ProtectionType.All example
// Common Searches: protect every sheet in an Excel file with Aspose.Cells | apply same password to all worksheets .NET | keep formatting when protecting Excel sheets programmatically | Aspose.Cells protect multiple worksheets example
// Developer Intent: Add identical password protection to every worksheet in a workbook without modifying any existing formatting.
// Use Cases: Lock all tabs of a financial report before sending to clients while preserving the visual layout. | Automate organization‑wide policy that enforces sheet protection on generated workbooks. | Create a template where users can view formatting but cannot edit cell contents on any sheet.
// AI Prompts: Write C# code with Aspose.Cells that protects all worksheets in a workbook using a single password and retains all formatting. | Show how to add robust error handling around worksheet protection for multiple Excel files. | Modify the example to protect only formulas and objects while leaving cell editing allowed.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, loops through every Worksheet, applies ProtectionType.All with a shared password, and saves the file while keeping all cell, row, column, and style formatting intact. Includes file‑existence check and exception handling.
    public class ProtectAllWorksheets
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Define the password to be applied to every worksheet
            const string password = "SecurePassword123";

            // Input and output file paths
            const string inputPath = "input.xlsx";
            const string outputPath = "output_protected.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet and protect it with all protection types
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The third parameter (oldPassword) is set to null because the sheets are not previously password‑protected
                    sheet.Protect(ProtectionType.All, password, null);
                }

                // Save the protected workbook; formatting of cells, rows, columns, etc., remains unchanged
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions (e.g., Aspose.Cells errors, IO issues)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
