// Title: Apply distinct passwords to each worksheet in an Excel file using Aspose.Cells for .NET based on the workbook name
// AI Prompts: Generate C# code with Aspose.Cells that protects every worksheet using a password formed from the workbook filename and the sheet name. | Show how to save the password‑protected workbook to a new location, automatically creating the output folder when it does not exist. | Add error handling that verifies the source file exists and reports any failures during worksheet protection or file saving.
// Common Searches: Aspose.Cells C# protect each worksheet with a different password derived from file name | How to set unique passwords for multiple sheets in an Excel workbook using .NET | Save a protected Excel workbook to a new file and ensure output directory exists in C# | C# code to apply per‑sheet protection with Aspose.Cells and handle missing input file
// Tags: Aspose.Cells per‑worksheet password protection | generate worksheet password from workbook filename | save protected Excel workbook Aspose.Cells .NET | C# create output directory if missing Aspose.Cells | error handling missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, derives a base password from the file name, and iterates through all worksheets, protecting each with a unique password that combines the base name and the sheet name using ProtectionType.All. It ensures the output directory exists, saves the protected workbook to a new file, and includes checks for a missing source file and general exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = @"C:\Data\SampleWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Derive a base password from the file name (without extension)
            string basePassword = Path.GetFileNameWithoutExtension(inputPath);

            // Protect each worksheet with a unique password
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a unique password for the current sheet
                string sheetPassword = $"{basePassword}_{sheet.Name}";

                // Apply protection with the generated password.
                // The third parameter (oldPassword) is not required for new protection, so pass null.
                sheet.Protect(ProtectionType.All, sheetPassword, null);
            }

            // Path for the protected workbook
            string outputPath = @"C:\Data\SampleWorkbook_Protected.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
