// Title: C# example: Unprotect a password‑protected worksheet in an XLSX workbook using Aspose.Cells
// Common Searches: how to remove worksheet password in an XLSX file with Aspose.Cells C# | Aspose.Cells .NET code to unprotect a specific sheet after loading a protected workbook | load an Excel workbook with a workbook password and then unprotect a sheet using Aspose.Cells | C# example for worksheet.Unprotect method in Aspose.Cells | unprotect a locked Excel worksheet without losing workbook data using Aspose.Cells
// Tags: Aspose.Cells worksheet unprotect C# | remove Excel sheet password Aspose.Cells | load password‑protected XLSX with LoadOptions Aspose.Cells | unprotect specific worksheet by name Aspose.Cells | C# Aspose.Cells worksheet protection removal | Aspose.Cells workbook and sheet password handling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an XLSX workbook (optionally using a workbook password via LoadOptions), accesses a target worksheet by name or index, calls Worksheet.Unprotect with the worksheet password, and saves the unprotected workbook to a new file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string workbookPassword = "";      // Set workbook password if needed
            const string worksheetPassword = "YourPassword"; // Worksheet password

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook with password (if any)
                LoadOptions loadOptions = new LoadOptions();
                if (!string.IsNullOrEmpty(workbookPassword))
                {
                    loadOptions.Password = workbookPassword;
                }

                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Access the target worksheet (by name or fallback to first sheet)
                Worksheet worksheet = workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    worksheet = workbook.Worksheets[0];
                }

                // Unprotect the worksheet
                worksheet.Unprotect(worksheetPassword);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
