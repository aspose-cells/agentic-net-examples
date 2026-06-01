using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (columns A to J)
            // Header row
            cells["A1"].PutValue("Group");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Qty");
            cells["D1"].PutValue("Price");
            cells["E1"].PutValue("Discount");
            cells["F1"].PutValue("Tax");
            cells["G1"].PutValue("Cost");
            cells["H1"].PutValue("Revenue");
            cells["I1"].PutValue("Category");
            cells["J1"].PutValue("Factor"); // Column J will receive the Product subtotal

            // Sample data rows (10 rows)
            string[] groups = { "North", "North", "South", "South", "East", "East", "West", "West", "North", "South" };
            Random rnd = new Random();
            for (int i = 0; i < 10; i++)
            {
                int row = i + 1; // zero‑based index for Cells is row number
                cells[row, 0].PutValue(groups[i]);               // Group column (A)
                cells[row, 1].PutValue($"Item{i + 1}");          // Item column (B)
                cells[row, 2].PutValue(rnd.Next(1, 10));         // Qty (C)
                cells[row, 3].PutValue(rnd.Next(10, 100));       // Price (D)
                cells[row, 4].PutValue(rnd.NextDouble());        // Discount (E)
                cells[row, 5].PutValue(rnd.NextDouble());        // Tax (F)
                cells[row, 6].PutValue(rnd.Next(5, 20));         // Cost (G)
                cells[row, 7].PutValue(rnd.Next(20, 200));       // Revenue (H)
                cells[row, 8].PutValue($"Cat{(i % 3) + 1}");      // Category (I)
                cells[row, 9].PutValue(rnd.Next(2, 5));          // Factor (J) – values to be multiplied
            }

            // Define the range that contains the data (including header)
            // From A1 (row 0, col 0) to J11 (row 10, col 9)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 10,
                EndColumn = 9
            };

            // Apply subtotals:
            // - Group by column A (index 0) – the "Group" field
            // - Use Product function on column J (index 9)
            // - Place summary rows above the detail rows (summaryBelowData = false)
            // - Do not replace existing subtotals, no page breaks
            cells.Subtotal(
                area,
                0,                                 // groupBy column index (A)
                ConsolidationFunction.Product,     // Product function
                new int[] { 9 },                   // subtotal on column J
                false,                             // replace existing subtotals
                false,                             // page breaks between groups
                false);                            // summaryBelowData = false (top)

            // Save the workbook
            workbook.Save("Subtotal_Product_Top.xlsx");
        }
    }
}