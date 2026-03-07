using System;
using Aspose.Cells;

namespace AsposeCellsShowHideDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (5 rows x 5 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ----- Hide rows and columns -----
            // Hide rows 2 to 4 (zero‑based indices 1,2,3)
            cells.HideRows(1, 3);

            // Hide column B (zero‑based index 1)
            cells.HideColumn(1);

            // Save the workbook after hiding
            workbook.Save("HiddenRowsColumns.xlsx", SaveFormat.Xlsx);

            // ----- Show (unhide) rows and columns -----
            // Unhide the previously hidden rows; height = -1 means auto‑fit
            cells.UnhideRows(1, 3, -1);

            // Unhide the previously hidden column; use standard column width
            cells.UnhideColumn(1, cells.StandardWidth);

            // Save the workbook after unhiding
            workbook.Save("UnhiddenRowsColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}