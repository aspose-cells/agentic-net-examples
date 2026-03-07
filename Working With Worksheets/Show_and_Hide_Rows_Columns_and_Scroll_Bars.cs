using System;
using Aspose.Cells;

class ShowHideDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill the worksheet with sample data (20 rows × 5 columns)
        for (int r = 0; r < 20; r++)
        {
            for (int c = 0; c < 5; c++)
            {
                cells[r, c].PutValue($"R{r + 1}C{c + 1}");
            }
        }

        // Hide a single row (5th row, zero‑based index 4)
        cells.HideRow(4);

        // Hide multiple rows (rows 10‑12, zero‑based start index 9, count 3)
        cells.HideRows(9, 3);

        // Hide a single column (column C, zero‑based index 2)
        cells.HideColumn(2);

        // Hide multiple columns (columns D‑E, start index 3, count 2)
        cells.HideColumns(3, 2);

        // Set the first visible row to row 8 (zero‑based index 7)
        sheet.FirstVisibleRow = 7;

        // Hide both horizontal and vertical scroll bars
        workbook.Settings.IsHScrollBarVisible = false;
        workbook.Settings.IsVScrollBarVisible = false;

        // Save the workbook to an XLSX file
        workbook.Save("ShowHideDemo.xlsx", SaveFormat.Xlsx);
    }
}