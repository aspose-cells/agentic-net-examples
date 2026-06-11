using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data before insertion (rows 1‑12)
            for (int i = 0; i < 12; i++)
            {
                cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Insert five rows at position 10 (zero‑based index 9)
            cells.InsertRows(9, 5);

            // Add data to the newly inserted rows (optional)
            for (int i = 9; i < 14; i++)
            {
                cells[i, 1].PutValue($"Inserted {i - 8}");
            }

            // Save the workbook
            workbook.Save("InsertFiveRowsAtTen.xlsx");
        }
    }
}