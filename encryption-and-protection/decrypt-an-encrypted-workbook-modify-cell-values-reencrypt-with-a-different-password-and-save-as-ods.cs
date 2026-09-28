// Title: Decrypt an encrypted XLSX workbook, change cell A1, re‑encrypt with a new password, and export to ODS using Aspose.Cells for .NET
// AI Prompts: Load a password‑protected XLSX file with the original password, set cell A1 to "Hello World", assign a new workbook password, and save the file as ODS using Aspose.Cells in C#. | Open an encrypted Excel workbook via LoadOptions, update a specific cell, change the workbook's encryption password, and export the result to OpenDocument Spreadsheet format with Aspose.Cells for .NET.
// Common Searches: How to open an encrypted XLSX in C# with Aspose.Cells and change its password | Aspose.Cells example for modifying a cell in a password‑protected workbook and saving as ODS | Re‑encrypt Excel file with a new password and convert to ODS using Aspose.Cells .NET | LoadOptions password property usage for encrypted Excel files in Aspose.Cells
// Tags: load encrypted xlsx with LoadOptions Aspose.Cells | modify cell value in workbook C# | set new workbook password Aspose.Cells | save workbook as ods format | re‑encrypt Excel file using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates loading an encrypted XLSX workbook with the original password, updating cell A1, applying a new password, and saving the workbook as an ODS file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Paths and passwords
            string inputFile = @"C:\Temp\EncryptedWorkbook.xlsx";
            string outputFile = @"C:\Temp\ReEncryptedWorkbook.ods";
            string oldPassword = "oldPassword123";
            string newPassword = "newPassword456";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the encrypted workbook using the old password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Example modification: set A1 to "Hello World"
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Hello World");

                // Apply new password (encryption algorithm property not available in this version)
                workbook.Settings.Password = newPassword;

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save as ODS (use the non‑obsolete enum value)
                workbook.Save(outputFile, SaveFormat.Ods);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
