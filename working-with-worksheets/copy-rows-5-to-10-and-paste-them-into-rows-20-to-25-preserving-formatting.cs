using System;
using System.Drawing;
using Aspose.Cells;

class CopyRowsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill rows 5 to 10 (zero‑based indices 4‑9) with sample data and a simple style
        for (int row = 4; row <= 9; row++)
        {
            cells[row, 0].PutValue($"Row {row + 1} Column A");
            cells[row, 1].PutValue(row * 10);

            // Apply formatting to demonstrate that formatting is preserved during copy
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Blue;
            style.Font.IsBold = true;
            cells[row, 0].SetStyle(style);
        }

        // Copy the six rows (5‑10) to rows 20‑25 (indices 19‑24) preserving data and formatting
        // Parameters: sourceCells, sourceRowIndex, destinationRowIndex, rowNumber
        cells.CopyRows(cells, 4, 19, 6);

        // Save the workbook
        workbook.Save("CopyRowsResult.xlsx");
    }
}