// Title: Load a mixed‑encoding CSV with TxtLoadOptions.IsMultiEncoded and preserve Unicode characters when converting to XLSX using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load CSV with multiple encodings using TxtLoadOptions.IsMultiEncoded | preserve Unicode characters from a mixed‑encoding CSV when converting to XLSX with Aspose.Cells | enable multi‑encoding detection for CSV files in Aspose.Cells .NET API | read CSV containing UTF‑8 and UTF‑16 data in Aspose.Cells without losing characters | convert mixed‑encoding CSV to Excel workbook using Aspose.Cells and verify cell values
// Tags: TxtLoadOptions.IsMultiEncoded CSV import | multi‑encoding CSV loading Aspose.Cells | Unicode preservation CSV to XLSX | Aspose.Cells mixed encoding detection | load CSV with multiple encodings .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for a mixed‑encoding CSV file, enables TxtLoadOptions.IsMultiEncoded to detect multiple encodings, loads the CSV into an Aspose.Cells Workbook, reads a Unicode value from cell A1, prints it, and saves the workbook as an XLSX file, handling any exceptions that may occur.
    class Program
    {
        static void Main()
        {
            // Path to the mixed‑encoding CSV file.
            string csvPath = "mixed_encoding.csv";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Input file not found: {csvPath}");
                return;
            }

            try
            {
                // Enable multi‑encoding detection for the CSV.
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    IsMultiEncoded = true
                };

                // Load the CSV file with the specified options.
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Access the first worksheet (CSV is loaded into a single sheet).
                Worksheet sheet = workbook.Worksheets[0];

                // Read a cell that may contain Unicode characters.
                string unicodeValue = sheet.Cells["A1"].StringValue;
                Console.WriteLine("Value in A1: " + unicodeValue);

                // Save the workbook to another format (e.g., XLSX) to confirm data integrity.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
