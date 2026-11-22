// Title: How to protect each worksheet with a unique password derived from the workbook filename using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells protect each worksheet with a different password derived from file name | set individual sheet passwords in an Excel file using Aspose.Cells .NET | apply full protection to all worksheets with unique passwords in Aspose.Cells | generate worksheet password from workbook name Aspose.Cells C# example
// Tags: Aspose.Cells worksheet protection per sheet | C# generate unique sheet passwords from filename | Aspose.Cells Protect method with ProtectionType.All | Excel workbook password derivation using file name | save protected workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // The example loads an existing Excel workbook, uses the file name (without extension) as a base password, creates a distinct password for each worksheet by appending the sheet name, applies full protection to every sheet via the Protect method with ProtectionType.All, and saves the protected workbook to a new file while handling missing files and directory creation.
    class WorksheetProtection
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string inputPath = @"C:\Data\SampleWorkbook.xlsx";

                // Path where the protected workbook will be saved
                string outputPath = @"C:\Data\SampleWorkbook_Protected.xlsx";

                // Verify input file exists
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
                    // Create a unique password per sheet (e.g., "SampleWorkbook_Sheet1")
                    string sheetPassword = $"{basePassword}_{sheet.Name}";

                    // Apply protection using the derived password.
                    // The third parameter (oldPassword) is not required for new protection, so pass empty string.
                    sheet.Protect(ProtectionType.All, sheetPassword, string.Empty);
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Protected workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
