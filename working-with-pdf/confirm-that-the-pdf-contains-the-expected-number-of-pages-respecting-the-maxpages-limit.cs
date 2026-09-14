// Title: Validate Excel workbook worksheet count against a maximum page limit using Aspose.Cells in C#
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells, counts its worksheets, and returns true if the count does not exceed a specified maxPages value. | Create a console application that accepts an Excel file path and a max page limit as command‑line arguments, invokes the worksheet‑count validator, and outputs a pass/fail message. | Enhance the validator to log the actual worksheet count and throw a custom MaxPagesExceededException when the limit is surpassed.
// Common Searches: how to enforce a maximum number of worksheets before converting Excel to PDF with Aspose.Cells C# | C# code to check Excel workbook page count using Aspose.Cells | validate worksheet count against limit Aspose.Cells .NET console app | Aspose.Cells verify number of sheets does not exceed max pages | programmatically restrict Excel sheets count for PDF export in C#
// Tags: Aspose.Cells worksheet count validation | C# max worksheets check for PDF export | Excel workbook page limit enforcement Aspose | validate worksheet number before PDF conversion .NET | custom exception for max pages exceeded Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace PdfPageValidationApp
{
    // The example provides a PdfPageValidator class with a CheckPdfPageCount method that loads an Excel workbook via Aspose.Cells, counts its worksheets, compares the count to a supplied maxPages limit, and returns a boolean result while handling missing files and exceptions. A console Program parses the Excel file path and limit from command‑line arguments, calls the validator, and prints a pass/fail message.
    public class PdfPageValidator
    {
        /// <param name="excelPath">Full file path to the Excel document.</param>
        /// <param name="maxPages">Maximum allowed number of worksheets.</param>
        /// <returns>True if the worksheet count is less than or equal to maxPages; otherwise false.</returns>
        public static bool CheckPdfPageCount(string excelPath, int maxPages)
        {
            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"File not found: {excelPath}");

            try
            {
                // Load the Excel workbook using Aspose.Cells
                Workbook workbook = new Workbook(excelPath);

                // Retrieve the total number of worksheets (treated as pages)
                int pageCount = workbook.Worksheets.Count;

                // Compare the count with the allowed maximum
                return pageCount <= maxPages;
            }
            catch (Exception ex)
            {
                // Log the exception and treat as validation failure
                Console.Error.WriteLine($"Error validating workbook: {ex.Message}");
                return false;
            }
        }
    }

    public class Program
    {
        // Entry point of the console application
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: PdfPageValidator <ExcelFilePath> <MaxPages>");
                    return;
                }

                string excelPath = args[0];
                if (!int.TryParse(args[1], out int maxPages))
                {
                    Console.WriteLine("Invalid max pages value.");
                    return;
                }

                bool result = PdfPageValidator.CheckPdfPageCount(excelPath, maxPages);
                Console.WriteLine(result
                    ? "Validation passed: worksheet count is within the allowed limit."
                    : "Validation failed: worksheet count exceeds the allowed limit.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions to prevent the application from crashing
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
