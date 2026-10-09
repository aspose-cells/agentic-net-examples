// Title: Check Excel workbook worksheet count against a maximum limit using Aspose.Cells in a C# console app
// AI Prompts: Create a C# method that loads an Excel file with Aspose.Cells, counts its worksheets, and throws an InvalidOperationException when the count exceeds a supplied maxSheets value. | Write a C# console program that accepts a file path and a maximum sheet count, calls the worksheet‑validation method, and logs either a success message or the caught exception.
// Common Searches: aspnet aspose.cells enforce maximum number of worksheets in an Excel file | c# console application validate workbook sheet count does not exceed limit | how to throw error if Excel workbook contains more sheets than allowed using Aspose.Cells
// Tags: Aspose.Cells worksheet count validation | C# enforce max worksheets | Excel workbook sheet limit check | console app workbook validation | InvalidOperationException on excess sheets

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example defines a PdfPageValidator (misnamed) class that uses Aspose.Cells to load an Excel workbook, compares the total number of worksheets to a user‑provided maximum, throws an InvalidOperationException if the limit is exceeded, and writes a confirmation message when the count is within bounds.
    class PdfPageValidator
    {
        // Validates that the Excel file at the given path does not exceed the specified maximum number of worksheets.
        public static void Validate(string excelPath, int maxSheets)
        {
            // Ensure the file exists to avoid FileNotFoundException.
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"The file \"{excelPath}\" was not found.");

            try
            {
                // Load the Excel workbook.
                Workbook workbook = new Workbook(excelPath);

                // Retrieve the total number of worksheets in the workbook.
                int sheetCount = workbook.Worksheets.Count;

                // Compare the sheet count with the allowed maximum.
                if (sheetCount > maxSheets)
                {
                    // Throw an exception if the workbook exceeds the limit.
                    throw new InvalidOperationException(
                        $"Workbook contains {sheetCount} worksheets, which exceeds the maximum allowed {maxSheets} worksheets.");
                }

                // If within limit, optionally inform the caller.
                Console.WriteLine($"Workbook worksheet count ({sheetCount}) is within the allowed limit of {maxSheets} worksheets.");
            }
            catch (Exception ex) when (!(ex is FileNotFoundException))
            {
                // Wrap any unexpected exceptions with additional context.
                throw new InvalidOperationException($"Failed to validate the workbook at \"{excelPath}\": {ex.Message}", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide excel file path and max sheets via command line or defaults.
            string excelPath = args.Length > 0 ? args[0] : "sample.xlsx";
            int maxSheets = (args.Length > 1 && int.TryParse(args[1], out int parsed)) ? parsed : 5;

            try
            {
                PdfPageValidator.Validate(excelPath, maxSheets);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
