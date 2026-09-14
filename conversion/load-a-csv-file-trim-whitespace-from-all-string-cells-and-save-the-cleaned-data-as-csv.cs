// Title: Trim whitespace from all string cells in a CSV using Aspose.Cells for .NET and save the cleaned file
// AI Prompts: Load a CSV file with Aspose.Cells, iterate through each cell, replace string values with their Trim() result, and write the workbook to a new CSV. | Configure TxtSaveOptions to drop empty leading rows/columns and trailing blank cells while exporting the cleaned worksheet as CSV.
// Common Searches: Aspose.Cells .NET how to remove spaces from string values in a CSV workbook | C# trim whitespace in all cells when processing CSV with Aspose.Cells | Save CSV with Aspose.Cells while eliminating blank rows and trailing empty cells | Iterate over cells in Aspose.Cells to clean data before exporting to CSV | Load CSV using LoadOptions and clean string fields in C#
// Tags: remove surrounding spaces from CSV cells Aspose.Cells | LoadOptions CSV import Aspose.Cells .NET | TxtSaveOptions eliminate empty rows columns CSV | cell iteration data sanitization Aspose.Cells | save cleaned worksheet as CSV Aspose.Cells

using System;
using Aspose.Cells;

namespace CsvTrimExample
{
    // The program loads an input CSV with Aspose.Cells, walks through every used cell, trims leading and trailing whitespace from string cells, and saves the cleaned worksheet to a new CSV using TxtSaveOptions that also discard empty leading rows/columns and trailing blank cells.
    class Program
    {
        static void Main()
        {
            // Paths for input and output CSV files
            string inputCsvPath = "input.csv";
            string outputCsvPath = "output.csv";

            // Load the CSV file using LoadOptions with CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputCsvPath, loadOptions);

            // Get the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Iterate through all used cells and trim whitespace from string values
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    // Trim leading and trailing whitespace and put the trimmed value back
                    string trimmed = cell.StringValue.Trim();
                    cell.PutValue(trimmed);
                }
            }

            // Prepare CSV save options (default separator is comma)
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Ensure leading blank rows/columns are trimmed (optional)
                TrimLeadingBlankRowAndColumn = true,
                // Ensure trailing blank cells are trimmed (optional)
                TrimTailingBlankCells = true
            };

            // Save the cleaned data back to CSV
            workbook.Save(outputCsvPath, saveOptions);
        }
    }
}
