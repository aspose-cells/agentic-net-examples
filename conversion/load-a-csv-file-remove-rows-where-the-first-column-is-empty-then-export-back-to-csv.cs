// Title: C# – Remove rows with an empty first column from a CSV using Aspose.Cells and save the cleaned file
// Common Searches: Aspose.Cells C# delete rows where first column is blank in a CSV file | How to filter out empty rows from a CSV using Aspose.Cells .NET | C# code to load a CSV, remove rows with a null first column, and export back to CSV with Aspose.Cells | Saving a modified CSV after row deletion with Aspose.Cells in .NET | Iterate rows backwards to safely delete rows in an Aspose.Cells workbook
// Tags: delete blank rows Aspose.Cells CSV | Aspose.Cells row deletion C# | load CSV workbook Aspose.Cells .NET | save workbook as CSV Aspose.Cells | iterate rows backwards Aspose.Cells

using System;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook, iterates the rows from the bottom up, deletes any row whose first column is null or whitespace, and then saves the cleaned data to a new CSV file.
class CsvProcessor
{
    static void Main()
    {
        // Load the CSV file into a workbook.
        // Aspose.Cells automatically treats the CSV as a single worksheet.
        Workbook workbook = new Workbook("input.csv");

        // Get the first (and only) worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Iterate from the last row upwards to safely delete rows.
        for (int row = cells.MaxDataRow; row >= cells.MinDataRow; row--)
        {
            // Retrieve the value of the first column (index 0) in the current row.
            object firstCellValue = cells[row, 0].Value;

            // Determine if the cell is empty (null or empty string after trimming).
            bool isEmpty = firstCellValue == null ||
                           string.IsNullOrWhiteSpace(firstCellValue.ToString());

            // If the first column is empty, delete the entire row.
            if (isEmpty)
            {
                cells.DeleteRow(row);
            }
        }

        // Save the modified worksheet back to CSV format.
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
