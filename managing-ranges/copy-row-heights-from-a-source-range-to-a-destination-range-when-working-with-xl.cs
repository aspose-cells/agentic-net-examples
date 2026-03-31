using System;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

class CopyRowHeights
{
    static void Main()
    {
        // Load the source workbook (contains the rows whose heights we want to copy)
        Workbook sourceWorkbook = new Workbook("source.xlsx");

        // Load the destination workbook (where the heights will be applied)
        Workbook destinationWorkbook = new Workbook("destination.xlsx");

        // Get the first worksheet from each workbook (adjust index if needed)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Define the source range whose row heights will be copied.
        // Example: rows 2 to 5 (Excel 1‑based indexing). Aspose.Cells uses zero‑based indices internally.
        string sourceRangeAddress = "2:5";
        CellsRange sourceRange = sourceSheet.Cells.CreateRange(sourceRangeAddress);

        // Define the first row index in the destination sheet where the heights will be applied.
        // Example: start copying to row 10 (zero‑based index = 9).
        int destinationStartRowIndex = 9; // corresponds to Excel row 10

        // Number of rows to copy
        int rowsToCopy = sourceRange.RowCount;

        // Iterate through each row in the source range and copy its settings (height, style, hidden flag, etc.)
        for (int i = 0; i < rowsToCopy; i++)
        {
            int sourceRowIndex = sourceRange.FirstRow + i;
            int destinationRowIndex = destinationStartRowIndex + i;

            Row sourceRow = sourceSheet.Cells.Rows[sourceRowIndex];
            Row destinationRow = destinationSheet.Cells.Rows[destinationRowIndex];

            // Copy settings from source row to destination row.
            // The second parameter (checkStyle) is false because both rows belong to the same workbook.
            destinationRow.CopySettings(sourceRow, false);
        }

        // Save the destination workbook with the updated row heights.
        destinationWorkbook.Save("result.xlsx");
    }
}