using System;
using Aspose.Cells;

namespace FreezeTopRowsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to illustrate the freeze)
            for (int i = 0; i < 20; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
                worksheet.Cells[i, 1].PutValue($"Data {i + 1}");
            }

            // Freeze the top three rows.
            // Parameters: row index, column index, number of frozen rows, number of frozen columns.
            // Row index 3 means the split occurs just below row 3 (zero‑based), freezing rows 0‑2.
            // Column index 0 and frozenColumns 0 means no columns are frozen.
            worksheet.FreezePanes(3, 0, 3, 0);

            // Save the workbook to a file
            workbook.Save("FreezeTopThreeRows.xlsx");
        }
    }
}