using System;
using Aspose.Cells;

namespace ConcealRowsColumnsDemo
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

            // Conceal rows 2 through 4 (zero‑based index 1 to 3)
            int startRow = 1;      // Row 2
            int rowCount = 3;      // Rows 2,3,4
            cells.HideRows(startRow, rowCount);

            // Conceal columns B through D (zero‑based index 1 to 3)
            int startColumn = 1;   // Column B
            int columnCount = 3;   // Columns B,C,D
            cells.HideColumns(startColumn, columnCount);

            // Save the workbook
            workbook.Save("ConcealedRowsColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}