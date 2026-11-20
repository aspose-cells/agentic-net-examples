// Title: Delete rows with an empty first column from a CSV using Aspose.Cells for .NET and save the result
// Common Searches: Aspose.Cells C# delete rows where the first column is blank in a CSV file | How to filter out empty‑first‑column rows from a CSV using Aspose.Cells .NET | C# code to remove rows with null first column from a CSV workbook with Aspose.Cells | Save a cleaned CSV after removing rows with empty first column using Aspose.Cells | Iterate rows backwards to delete rows in Aspose.Cells CSV processing
// Tags: remove rows Aspose.Cells CSV | delete rows based on first column Aspose.Cells | save workbook as CSV Aspose.Cells | reverse row iteration Aspose.Cells | trim cell values Aspose.Cells CSV

using System;
using Aspose.Cells;

// The example loads 'input.csv' into an Aspose.Cells Workbook, iterates rows from bottom to top, removes any row whose first column is null or empty, and then saves the cleaned data back to 'output.csv' in CSV format.
class CsvProcessor
{
    static void Main()
    {
        // Load the CSV file into a workbook
        Workbook workbook = new Workbook("input.csv");

        // Get the first worksheet (CSV files are loaded into a single sheet)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Iterate from the last row to the first to safely remove rows
        for (int rowIndex = cells.MaxDataRow; rowIndex >= cells.MinDataRow; rowIndex--)
        {
            // Get the value of the first column (index 0) in the current row
            Cell firstCell = cells[rowIndex, 0];
            string cellValue = firstCell?.StringValue?.Trim();

            // If the first column is empty or null, remove the entire row
            if (string.IsNullOrEmpty(cellValue))
            {
                cells.Rows.RemoveAt(rowIndex);
            }
        }

        // Save the modified worksheet back to CSV format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
