using System;
using Aspose.Cells;

namespace SubtotalAverageColumnF
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // Sample data setup (columns A to F, rows 1 to 10)
            // -------------------------------------------------
            // Header row
            cells["A1"].PutValue("Group");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Qty");
            cells["D1"].PutValue("Price");
            cells["E1"].PutValue("Date");
            cells["F1"].PutValue("Score"); // Column F (index 5) – the column we will average

            // Populate some rows with dummy data
            for (int row = 1; row <= 9; row++)
            {
                cells[row, 0].PutValue(row % 2 == 0 ? "B" : "A");          // Group
                cells[row, 1].PutValue($"Item{row}");                     // Item
                cells[row, 2].PutValue(row * 10);                         // Qty
                cells[row, 3].PutValue(row * 5.5);                        // Price
                cells[row, 4].PutValue(DateTime.Today.AddDays(-row));    // Date
                cells[row, 5].PutValue(row * 2);                          // Score (numeric)
            }

            // -------------------------------------------------
            // Define the range that contains the data (A1:F10)
            // -------------------------------------------------
            CellArea dataArea = CellArea.CreateCellArea("A1", "F10");

            // -------------------------------------------------
            // Add subtotals:
            // - Group by the first column (index 0)
            // - Use Average function
            // - Apply to column F (index 5)
            // - Replace existing subtotals (true)
            // - No page breaks between groups (false)
            // - Place summary rows above the data (summaryBelowData = false)
            // -------------------------------------------------
            cells.Subtotal(
                dataArea,
                0,                                 // groupBy column (A)
                ConsolidationFunction.Average,    // average calculation
                new int[] { 5 },                  // subtotal on column F
                true,                             // replace existing subtotals
                false,                            // no page breaks
                false);                           // summary rows at the top

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("Subtotal_Average_ColumnF.xlsx");
        }
    }
}