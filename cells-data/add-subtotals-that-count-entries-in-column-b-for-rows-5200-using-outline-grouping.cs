using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column B (index 1) for rows 5‑200 (Excel rows 5‑200 correspond to zero‑based indices 4‑199)
            for (int row = 4; row <= 199; row++)
            {
                // Example: put a category name that repeats to demonstrate grouping
                string value = (row % 10 == 0) ? "GroupA" : "GroupB";
                cells[row, 1].PutValue(value);
            }

            // Define the cell area that contains the data to subtotal (only column B)
            CellArea area = new CellArea
            {
                StartRow = 4,      // Row 5
                EndRow = 199,      // Row 200
                StartColumn = 1,   // Column B
                EndColumn = 1      // Column B
            };

            // Apply subtotal:
            // - groupBy: column B (index 1)
            // - function: Count (counts entries)
            // - totalList: apply the count to column B (index 1)
            cells.Subtotal(area, 1, ConsolidationFunction.Count, new int[] { 1 });

            // Ensure the summary row appears below the detail rows in the outline
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalOutline.xlsx");
        }
    }
}