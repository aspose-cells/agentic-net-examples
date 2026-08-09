// Title: Trim whitespace from all string cells in a CSV with Aspose.Cells for .NET
// Description: Load a CSV file using Aspose.Cells' TxtLoadOptions, iterate through every used cell, trim leading and trailing spaces from string values, update changed cells, and save the cleaned data back to CSV with TxtSaveOptions (comma separator, optional blank‑row/column trimming).
// Keywords: Aspose.Cells CSV trim | remove whitespace CSV .NET | trim string cells Aspose.Cells | CsvTrimExample C# | TxtLoadOptions TxtSaveOptions | clean CSV Aspose.Cells | trim leading blank rows columns CSV
// Common Searches: how to trim spaces from CSV cells using Aspose.Cells | Aspose.Cells remove whitespace from string values in CSV | save cleaned CSV after trimming cells Aspose.Cells .NET | trim all string cells in a CSV file C# | Aspose.Cells CSV whitespace cleanup example
// Developer Intent: Remove leading and trailing spaces from every string cell in a CSV file and write the sanitized content to a new CSV.
// Use Cases: Sanitize user‑uploaded CSVs before bulk import to prevent mismatched keys. | Prepare legacy CSV exports for reporting tools that are sensitive to extra spaces. | Normalize data for downstream analytics pipelines that require consistent string values.
// AI Prompts: Create C# code that logs each cell's original and trimmed value when a change occurs. | Show how to modify the example to trim only leading spaces while keeping trailing spaces. | Provide a version that processes all worksheets in the workbook instead of just the first one.

using System;
using Aspose.Cells;

namespace CsvTrimExample
{
    // Load a CSV file using Aspose.Cells' TxtLoadOptions, iterate through every used cell, trim leading and trailing spaces from string values, update changed cells, and save the cleaned data back to CSV with TxtSaveOptions (comma separator, optional blank‑row/column trimming).
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output CSV file paths
            string inputCsvPath = "input.csv";
            string outputCsvPath = "output_trimmed.csv";

            // Load CSV with default TxtLoadOptions (separator is comma by default)
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            Workbook workbook = new Workbook(inputCsvPath, loadOptions);

            // Get the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Iterate over all used cells and trim whitespace from string values
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    string original = cell.StringValue;
                    string trimmed = original.Trim();

                    // Update the cell only if trimming changed the value
                    if (!string.Equals(original, trimmed, StringComparison.Ordinal))
                    {
                        cell.PutValue(trimmed);
                    }
                }
            }

            // Prepare save options for CSV output
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Ensure leading blank rows/columns are trimmed (optional)
                TrimLeadingBlankRowAndColumn = true,
                // Use comma as separator
                Separator = ','
            };

            // Save the cleaned workbook as CSV
            workbook.Save(outputCsvPath, saveOptions);

            Console.WriteLine("CSV file has been trimmed and saved to: " + outputCsvPath);
        }
    }
}
