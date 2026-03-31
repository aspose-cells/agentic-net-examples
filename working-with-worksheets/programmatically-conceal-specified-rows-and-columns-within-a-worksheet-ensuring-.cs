using System;
using Aspose.Cells;

class HideRowsColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (10 rows x 8 columns)
        for (int i = 0; i < 10; i++)
        {
            for (int j = 0; j < 8; j++)
            {
                cells[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Hide rows 3 to 5 (zero‑based indices 2, 3, 4)
        cells.HideRows(2, 3);

        // Hide columns B to D (zero‑based indices 1, 2, 3)
        cells.HideColumns(1, 3);

        // Verify hidden status of rows
        for (int r = 0; r < 10; r++)
        {
            Console.WriteLine($"Row {r} hidden: {cells.IsRowHidden(r)}");
        }

        // Verify hidden status of columns
        for (int c = 0; c < 8; c++)
        {
            Console.WriteLine($"Column {c} hidden: {cells.IsColumnHidden(c)}");
        }

        // Save the workbook preserving hidden rows/columns
        workbook.Save("HiddenRowsColumns.xlsx", SaveFormat.Xlsx);
    }
}