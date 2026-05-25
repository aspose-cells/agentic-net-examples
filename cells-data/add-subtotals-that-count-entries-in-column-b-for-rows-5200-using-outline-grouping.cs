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

            // (Optional) Populate sample data in column B for rows 5‑200
            // Here we just fill with some example values; in real scenarios the data may already exist.
            for (int row = 4; row <= 199; row++) // zero‑based rows 4‑199 correspond to Excel rows 5‑200
            {
                // Example: put a category name that repeats every 5 rows
                string value = "Category" + ((row - 4) / 5);
                cells[row, 1].PutValue(value); // Column B has index 1
            }

            // Define the cell area that covers rows 5‑200 in column B
            CellArea area = new CellArea
            {
                StartRow = 4,      // Row 5 (zero‑based)
                StartColumn = 1,   // Column B
                EndRow = 199,      // Row 200 (zero‑based)
                EndColumn = 1      // Column B
            };

            // Add subtotals:
            // - Group by column B (index 1)
            // - Use COUNT function to count entries in column B
            // - Apply subtotal to column B (index 1)
            // - Replace existing subtotals, no page breaks, summary rows placed below data
            worksheet.Cells.Subtotal(
                area,
                1,                                 // groupBy column B
                ConsolidationFunction.Count,       // count entries
                new int[] { 1 },                   // subtotal column B
                true,                              // replace existing subtotals
                false,                             // no page breaks between groups
                true                               // summary rows below the detail rows
            );

            // Ensure the outline summary rows are positioned below the detail rows
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalOutlineDemo.xlsx");
        }
    }
}