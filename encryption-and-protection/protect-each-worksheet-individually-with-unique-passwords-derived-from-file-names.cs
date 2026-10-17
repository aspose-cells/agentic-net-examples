// Title: Protect each worksheet with a unique password based on the workbook file name using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells protect individual worksheets with different passwords | How to generate worksheet passwords from Excel file name using Aspose.Cells | Apply ProtectionType.All to each sheet and save the protected workbook in .NET | Set per‑sheet password in Aspose.Cells based on sheet name and file name | Save a protected Excel workbook to a new file with Aspose.Cells C# example
// Tags: Aspose.Cells worksheet.Protect dynamic password | generate per‑sheet password from workbook filename C# | protect Excel worksheets individually Aspose.Cells | save protected workbook to new file .NET | ProtectionType.All usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionExample
{
    // The example loads an existing Excel workbook, derives a base password from the file name, creates a unique password for each worksheet by appending the sheet name, protects each sheet with ProtectionType.All, and saves the protected workbook to a new file, creating the output directory if necessary and handling errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputFile = @"C:\Data\SampleWorkbook.xlsx";
            string outputFile = @"C:\Data\SampleWorkbook_Protected.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Derive a base password from the file name (without extension)
                string basePassword = Path.GetFileNameWithoutExtension(inputFile);

                // Protect each worksheet with a unique password
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Create a unique password for the worksheet (BaseFileName_SheetName)
                        string sheetPassword = $"{basePassword}_{sheet.Name}";

                        // Protect the worksheet; oldPassword is null because the sheet is not previously protected
                        sheet.Protect(ProtectionType.All, sheetPassword, null);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to protect sheet '{sheet.Name}': {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir!))
                {
                    Directory.CreateDirectory(outputDir!);
                }

                // Save the protected workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
