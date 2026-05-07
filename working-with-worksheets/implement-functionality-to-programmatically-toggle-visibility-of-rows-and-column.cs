using System;
using Aspose.Cells;

namespace AsposeCellsToggleVisibility
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int r = 0; r < 10; r++)
            {
                for (int c = 0; c < 5; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }

            // Toggle visibility: hide rows 2‑4 (zero‑based index 1‑3)
            HideRows(sheet, 1, 3, true);
            // Toggle visibility: hide columns 2‑3 (zero‑based index 1‑2)
            HideColumns(sheet, 1, 2, true);

            // Later we can unhide them dynamically
            HideRows(sheet, 1, 3, false);
            HideColumns(sheet, 1, 2, false);

            // Save the workbook
            workbook.Save("ToggleVisibilityDemo.xlsx");
        }

        // Helper method to hide or unhide a range of rows
        static void HideRows(Worksheet sheet, int startRow, int rowCount, bool hide)
        {
            if (hide)
            {
                // Hide multiple rows at once
                sheet.Cells.HideRows(startRow, rowCount);
            }
            else
            {
                // Unhide each row individually; provide a default height (15 points)
                for (int i = 0; i < rowCount; i++)
                {
                    sheet.Cells.UnhideRow(startRow + i, 15);
                }
            }
        }

        // Helper method to hide or unhide a range of columns
        static void HideColumns(Worksheet sheet, int startColumn, int columnCount, bool hide)
        {
            if (hide)
            {
                // Hide multiple columns at once
                sheet.Cells.HideColumns(startColumn, columnCount);
            }
            else
            {
                // Unhide each column individually; provide a default width (8.43 characters)
                for (int i = 0; i < columnCount; i++)
                {
                    sheet.Cells.UnhideColumn(startColumn + i, 8.43);
                }
            }
        }
    }
}