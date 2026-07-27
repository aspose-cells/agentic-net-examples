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

            // Populate sample data (including header)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Category
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Sales
            }

            // Determine the last used row and column
            int lastRow = cells.MaxDataRow;          // zero‑based index of the last row containing data
            int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column containing data

            // Define the cell area covering the entire used range (including header)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = lastColumn
            };

            // Apply subtotal:
            // - Group by the first column (Category) -> groupBy = 0
            // - Use SUM function for subtotals
            // - Add subtotal for the Sales column (index 2)
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

            // Save the workbook
            workbook.Save("SubtotalFullRangeDemo.xlsx");
        }
    }
}