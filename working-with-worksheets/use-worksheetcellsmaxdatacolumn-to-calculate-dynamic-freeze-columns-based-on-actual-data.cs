using System;
using Aspose.Cells;

class FreezeColumnsDynamic
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data to demonstrate dynamic detection of the last data column
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["C1"].PutValue("Header3");
        cells["A2"].PutValue(10);
        cells["B2"].PutValue(20);
        cells["C2"].PutValue(30);
        // Additional data in a farther column
        cells["E1"].PutValue("Header5");
        cells["E2"].PutValue(50);

        // Get the maximum column index that contains data (zero‑based)
        int maxDataColumn = cells.MaxDataColumn; // Returns -1 if no data

        if (maxDataColumn >= 0)
        {
            // Freeze the top row and all columns up to the last data column
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // row and column specify the cell where the split occurs (zero‑based)
            int splitRow = 1; // Freeze the first row (row index 0)
            int splitColumn = maxDataColumn + 1; // Cell after the last data column
            int frozenRows = 1; // Number of rows to keep visible
            int frozenColumns = maxDataColumn + 1; // Number of columns to keep visible

            sheet.FreezePanes(splitRow, splitColumn, frozenRows, frozenColumns);
        }

        // Save the workbook with the applied freeze panes
        workbook.Save("DynamicFreezeColumns.xlsx");
    }
}