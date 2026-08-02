// Title: Clean CSV files by trimming whitespace from every string cell with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to load a CSV document using Aspose.Cells' TxtLoadOptions, iterate through all cells in the first worksheet, remove leading and trailing spaces from string values, and export the sanitized data back to CSV with the original delimiter. Optional settings also eliminate empty rows and columns during save.
// Keywords: Aspose.Cells | C# CSV trim | remove whitespace from CSV cells | TxtLoadOptions CSV | TxtSaveOptions delimiter | clean CSV data | strip spaces Aspose.Cells | blank row removal | data cleansing CSV .NET | US developers | European .NET community
// Common Searches: Aspose.Cells trim spaces in CSV C# | how to remove leading/trailing whitespace from CSV cells using Aspose | save cleaned CSV with original separator Aspose.Cells | delete empty rows and columns when exporting CSV Aspose | C# code to sanitize CSV data with Aspose.Cells
// Developer Intent: Load a CSV, purge whitespace from all textual cells, and write the cleaned content back to a new CSV file.
// Use Cases: Pre‑process user‑submitted CSV uploads to avoid mismatched keys in a database. | Generate consistent CSV reports where stray spaces disrupt sorting or filtering. | Standardize input files for ETL pipelines that require exact string matches.
// AI Prompts: Write C# code that uses Aspose.Cells to open a CSV, trim whitespace from each string cell, and save it preserving the original comma delimiter. | Explain how TxtSaveOptions properties TrimLeadingBlankRowAndColumn and TrimTailingBlankCells affect the final CSV output. | Provide performance tips for trimming cells in large CSV files with Aspose.Cells while minimizing memory usage.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvTrimExample
{
    // This example demonstrates how to load a CSV document using Aspose.Cells' TxtLoadOptions, iterate through all cells in the first worksheet, remove leading and trailing spaces from string values, and export the sanitized data back to CSV with the original delimiter. Optional settings also eliminate empty rows and columns during save.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string inputCsvPath = "input.csv";

            // Load the CSV file using default TxtLoadOptions (CSV format)
            TxtLoadOptions loadOptions = new TxtLoadOptions(); // uses CSV by default
            Workbook workbook = new Workbook(inputCsvPath, loadOptions);

            // Get the first worksheet (the CSV data is loaded into the first sheet)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all cells and trim whitespace from string values
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    // Trim leading and trailing whitespace and write back the trimmed value
                    string trimmed = cell.StringValue.Trim();
                    cell.PutValue(trimmed);
                }
            }

            // Prepare save options for CSV output
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Ensure the same separator as the original CSV (comma)
                Separator = ',',
                // Optional: trim leading blank rows/columns like Excel does
                TrimLeadingBlankRowAndColumn = true,
                // Optional: trim trailing blank cells in each row
                TrimTailingBlankCells = true
            };

            // Save the cleaned data back to a CSV file
            string outputCsvPath = "output_cleaned.csv";
            workbook.Save(outputCsvPath, saveOptions);

            Console.WriteLine("CSV file has been trimmed and saved to: " + outputCsvPath);
        }
    }
}
