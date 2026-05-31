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
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Amount");

            // Sample data rows
            object[,] data = new object[,]
            {
                { "Fruit", "Apple", 120 },
                { "Fruit", "Banana", 80 },
                { "Fruit", "Orange", 150 },
                { "Vegetable", "Carrot", 60 },
                { "Vegetable", "Potato", 90 },
                { "Vegetable", "Tomato", 110 }
            };

            // Populate the worksheet starting from row 2 (zero‑based index 1)
            for (int r = 0; r < data.GetLength(0); r++)
                for (int c = 0; c < data.GetLength(1); c++)
                    cells[r + 1, c].PutValue(data[r, c]);

            // Define the range that contains the data (including header)
            CellArea area = CellArea.CreateCellArea("A1", "C7");

            // Apply subtotals:
            // - Group by the first column (Category) -> index 0
            // - Use SUM function
            // - Subtotal the third column (Amount) -> index 2
            // - Replace existing subtotals: false
            // - Add page breaks between groups: false
            // - Place summary rows below the data: true (bottom)
            cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 2 },                   // columns to subtotal
                false,                             // replace existing subtotals
                false,                             // page breaks between groups
                true                               // summaryBelowData (bottom)
            );

            // Ensure the outline also positions summary rows below detail rows
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalBottomPosition.xlsx");
        }
    }
}