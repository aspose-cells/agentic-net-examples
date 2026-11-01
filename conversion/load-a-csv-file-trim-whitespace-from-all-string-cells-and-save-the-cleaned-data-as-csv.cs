// Title: Trim whitespace from all string cells in a CSV file using Aspose.Cells for .NET and save the cleaned CSV
// Common Searches: C# remove leading and trailing spaces from CSV cells using Aspose.Cells | how to clean whitespace in string values of a CSV workbook with Aspose.Cells .NET | load a CSV, trim all string cells, and export a new CSV using Aspose.Cells | Aspose.Cells example for trimming spaces in CSV data programmatically | save cleaned CSV after whitespace removal with Aspose.Cells in C#
// Tags: Csv cell extra space removal Aspose.Cells | LoadOptions for CSV format Aspose.Cells | Iterate worksheet cells and update string values C# | TxtSaveOptions to export cleaned CSV Aspose.Cells | Aspose.Cells data sanitization for CSV files

using System;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook using LoadOptions, iterates through every used cell in the first worksheet, trims leading and trailing spaces from cells that contain string values, writes the trimmed text back to each cell, and then saves the modified worksheet as a new CSV file with TxtSaveOptions.
class CsvCleaner
{
    static void Main()
    {
        // Load the CSV file into a Workbook
        // LoadOptions can be used to specify CSV format if needed
        var loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook("input.csv", loadOptions);

        // Get the first worksheet (CSV data is usually in the first sheet)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Iterate through all used cells and trim whitespace from string values
        foreach (Cell cell in cells)
        {
            if (cell.Type == CellValueType.IsString)
            {
                // Trim leading and trailing whitespace
                string trimmed = cell.StringValue.Trim();
                // Update the cell with the trimmed value
                cell.PutValue(trimmed);
            }
        }

        // Save the cleaned data back to a CSV file
        var saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        workbook.Save("output.csv", saveOptions);
    }
}
