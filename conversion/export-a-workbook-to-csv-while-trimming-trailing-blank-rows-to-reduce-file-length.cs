// Title: Trim trailing blank rows from an Excel workbook and export to CSV using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells remove trailing empty rows before saving as csv | how to delete rows after the last data row in each worksheet using aspose.cells | export excel to csv without extra blank rows in asp.net with aspose.cells | trim blank rows in workbook then save as csv using aspose.cells c# | aspose.cells c# save first worksheet as csv after removing empty rows
// Tags: clean empty rows aspose.cells c# | export worksheet to csv aspose.cells | delete rows after MaxDataRow aspose.cells | save workbook as csv without empty rows | iterate worksheets aspose.cells c# | remove extra rows excel aspose.cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet to locate the last populated row using MaxDataRow, deletes any rows that follow, and then saves the first worksheet as a CSV file, eliminating trailing blank rows.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Process each worksheet to remove trailing blank rows
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the last row that contains data
            int lastDataRow = sheet.Cells.MaxDataRow; // zero‑based index

            // If there are rows after the last data row, delete them
            int totalRows = sheet.Cells.MaxRow; // total rows in the sheet (zero‑based)
            int rowsToDelete = totalRows - lastDataRow;

            if (rowsToDelete > 0)
            {
                // Delete rows starting from the row after the last data row
                // Parameters: start row index (zero‑based), number of rows to delete
                sheet.Cells.DeleteRows(lastDataRow + 1, rowsToDelete);
            }
        }

        // Save the trimmed workbook as CSV (only the first worksheet is exported by default)
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
