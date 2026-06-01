using System;
using Aspose.Cells;

namespace SubtotalAverageTopDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: headers in row 1, data starts from row 2
            // Column F (index 5) will contain numeric values for which we want the average subtotal
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Qty");
            cells["D1"].PutValue("Price");
            cells["E1"].PutValue("Date");
            cells["F1"].PutValue("Score"); // Column to average

            // Populate some rows of data
            object[,] data = new object[,]
            {
                { "A", "Item1", 10, 5.5, DateTime.Today, 80 },
                { "A", "Item2", 7,  3.2, DateTime.Today, 90 },
                { "B", "Item3", 5,  6.0, DateTime.Today, 75 },
                { "B", "Item4", 12, 4.1, DateTime.Today, 85 },
                { "C", "Item5", 9,  2.8, DateTime.Today, 95 }
            };

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]); // +1 because row 0 is header
                }
            }

            // Define the cell area that includes the header row and all data rows
            // StartRow = 0 (header), StartColumn = 0 (A), EndRow = rows (data rows count) , EndColumn = 5 (F)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = rows,          // rows count (header + data)
                EndColumn = 5           // column F (zero‑based index)
            };

            // Apply subtotal:
            // - Group by the first column (Category) -> groupBy = 0
            // - Use Average function
            // - Subtotal column is column F (index 5)
            // - Replace existing subtotals = false
            // - Add page breaks between groups = false
            // - SummaryBelowData = false (places summary rows at the top)
            cells.Subtotal(
                area,
                0,
                ConsolidationFunction.Average,
                new int[] { 5 },
                false,
                false,
                false);

            // Save the workbook
            workbook.Save("Subtotal_Average_Top.xlsx");
        }
    }
}