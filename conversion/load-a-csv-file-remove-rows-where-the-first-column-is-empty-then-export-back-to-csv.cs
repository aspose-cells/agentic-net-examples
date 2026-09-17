// Title: Remove rows with an empty first column from a CSV file using Aspose.Cells for .NET and save the cleaned CSV
// AI Prompts: Load a CSV into an Aspose.Cells Workbook, delete every row where column A is blank, and write the cleaned data to a new CSV file. | Create C# code that iterates from the last row upward, removes rows with an empty first‑column value using Aspose.Cells, and then exports the worksheet back to CSV.
// Common Searches: Aspose.Cells C# delete rows where first column is empty in a CSV | How to filter out blank rows from a CSV using Aspose.Cells .NET | C# remove rows with empty column A from CSV with Aspose.Cells and save | Iterate worksheet rows bottom‑up to delete rows in Aspose.Cells | Save modified worksheet as CSV using Aspose.Cells SaveOptions
// Tags: Aspose.Cells delete rows by column value | Aspose.Cells CSV row filtering | C# remove blank rows from CSV with Aspose.Cells | Aspose.Cells bottom‑up row deletion | Aspose.Cells save workbook as CSV | Aspose.Cells load CSV with LoadOptions

using System;
using Aspose.Cells;

// The program loads "input.csv" with Aspose.Cells, scans rows from the bottom up, deletes any row whose first column is empty, and saves the cleaned data as "output.csv" in CSV format.
class Program
{
    static void Main()
    {
        // Load the CSV file into a workbook
        var loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook("input.csv", loadOptions);

        // Get the first worksheet (CSV files are loaded into a single sheet)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range to know how many rows to process
        int maxRow = cells.MaxDataRow; // zero‑based index of the last row with data

        // Iterate from the bottom up to safely delete rows
        for (int row = maxRow; row >= 0; row--)
        {
            // Get the value of the first column (column index 0)
            var cellValue = cells[row, 0].StringValue?.Trim();

            // If the first column is empty or null, delete the entire row
            if (string.IsNullOrEmpty(cellValue))
            {
                sheet.Cells.DeleteRows(row, 1);
            }
        }

        // Save the modified worksheet back to CSV format
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Csv);
        workbook.Save("output.csv", saveOptions);
    }
}
