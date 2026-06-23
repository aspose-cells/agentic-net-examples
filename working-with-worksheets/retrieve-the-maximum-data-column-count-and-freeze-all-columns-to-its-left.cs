using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ------------------------------------------------------------
        // Sample data – this section can be replaced with loading an existing file
        // ------------------------------------------------------------
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["C1"].PutValue("Header3");
        cells["A2"].PutValue(100);
        cells["B2"].PutValue(200);
        cells["C2"].PutValue(300);
        // ------------------------------------------------------------

        // Retrieve the maximum column index that contains data (0‑based)
        int maxDataColumn = cells.MaxDataColumn; // Returns -1 if worksheet is empty

        if (maxDataColumn >= 0)
        {
            // Freeze all columns to the left of the column after the last data column.
            // FreezePanes(row, column, freezedRows, freezedColumns)
            //   row = 0            -> no rows are frozen
            //   column = maxDataColumn + 1 -> first column that remains scrollable
            //   freezedRows = 0   -> no rows frozen
            //   freezedColumns = maxDataColumn + 1 -> freeze all columns up to maxDataColumn
            worksheet.FreezePanes(0, maxDataColumn + 1, 0, maxDataColumn + 1);
        }

        // Save the workbook
        workbook.Save("FrozenColumns.xlsx", SaveFormat.Xlsx);
    }
}