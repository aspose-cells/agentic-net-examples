using System;
using Aspose.Cells;

namespace SubtotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Sample data (optional). In a real scenario the data would already exist.
            // Header row
            cells["A1"].PutValue("Group");      // Column A – grouping field
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Quantity");
            cells["D1"].PutValue("Amount");     // Column D – values to subtotal

            // Populate rows 2‑100 with dummy data
            for (int row = 1; row < 100; row++)          // zero‑based index: row 1 = Excel row 2
            {
                // Alternate groups for demonstration (e.g., "A", "B", "C")
                string group = ((row - 1) / 10 % 3) switch
                {
                    0 => "A",
                    1 => "B",
                    _ => "C"
                };
                cells[row, 0].PutValue(group);          // Group column
                cells[row, 1].PutValue($"Item {row}");
                cells[row, 2].PutValue(row * 2);        // Quantity
                cells[row, 3].PutValue(row * 5.5);      // Amount (column D)
            }

            // ------------------------------------------------------------
            // Define the range that contains the data (including header)
            // StartRow = 0 (row 1), StartColumn = 0 (A), EndRow = 99 (row 100), EndColumn = 3 (D)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 99,
                EndColumn = 3
            };

            // Apply subtotals:
            // - Group by the first column (index 0, "Group")
            // - Use SUM function
            // - Subtotal column D (zero‑based index 3)
            // - Replace existing subtotals (false = keep existing, true = replace)
            // - No page breaks between groups
            // - Place summary rows below each group (true)
            cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 3 },                   // columns to subtotal (D)
                true,                              // replace existing subtotals
                false,                             // no page breaks
                true                               // summary below data
            );

            // ------------------------------------------------------------
            // Save the workbook
            workbook.Save("SubtotalResult.xlsx");
        }
    }
}