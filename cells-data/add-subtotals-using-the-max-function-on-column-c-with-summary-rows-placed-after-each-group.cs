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

            // Populate sample data (columns A, B, C)
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Value");

            // Data rows
            object[,] data = new object[,]
            {
                { "Group1", "ItemA", 120 },
                { "Group1", "ItemB", 150 },
                { "Group1", "ItemC", 130 },
                { "Group2", "ItemA", 200 },
                { "Group2", "ItemB", 180 },
                { "Group2", "ItemC", 210 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the range that contains the data (including header)
            // A1:C7 -> rows 0-6, columns 0-2
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0),   // last data row index + header
                EndColumn = 2
            };

            // Add subtotals:
            // - Group by the first column (Category) -> groupBy = 0
            // - Use Max function on column C (zero‑based index 2)
            // - Replace existing subtotals = true
            // - No page breaks between groups = false
            // - Summary rows placed below the detail rows = true
            cells.Subtotal(
                area,
                0,
                ConsolidationFunction.Max,
                new int[] { 2 },
                true,
                false,
                true);

            // Save the workbook
            workbook.Save("Subtotal_Max_ColumnC_AfterEachGroup.xlsx");
        }
    }
}