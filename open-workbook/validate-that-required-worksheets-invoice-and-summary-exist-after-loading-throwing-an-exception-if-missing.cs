using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace WorksheetValidationExample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file to be loaded
                string inputPath = "input.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Check for required worksheets "Invoice" and "Summary"
                bool hasInvoice = workbook.Worksheets.Any(ws => ws.Name.Equals("Invoice", StringComparison.OrdinalIgnoreCase));
                bool hasSummary = workbook.Worksheets.Any(ws => ws.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase));

                // Throw an exception if any required worksheet is missing
                if (!hasInvoice || !hasSummary)
                {
                    string missing = (!hasInvoice ? "Invoice" : "") +
                                     (!hasInvoice && !hasSummary ? " and " : "") +
                                     (!hasSummary ? "Summary" : "");
                    throw new InvalidOperationException($"Required worksheet(s) missing: {missing}");
                }

                // Optional: proceed with further processing or save the workbook
                // workbook.Save("output.xlsx");
                Console.WriteLine("Workbook validation succeeded.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}