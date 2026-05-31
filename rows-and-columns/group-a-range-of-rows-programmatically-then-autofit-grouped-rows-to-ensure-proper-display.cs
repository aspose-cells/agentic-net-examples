using System;
using Aspose.Cells;

namespace AsposeCellsGroupRowsAutoFit
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in rows 1‑5 (zero‑based indices 0‑4)
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue($"Item {row + 1}");
                cells[row, 1].PutValue($"Description for item {row + 1} which may be long enough to require row height adjustment.");
            }

            // Group rows 2‑5 (zero‑based indices 1‑4) and hide them initially
            cells.GroupRows(1, 4, true);

            // Auto‑fit the grouped rows to adjust their heights based on the content
            worksheet.AutoFitRows(1, 4);

            // Save the workbook
            workbook.Save("GroupedRowsAutoFit.xlsx");
        }
    }
}