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

            // Populate sample data (header + 5 data rows)
            // Header: Region (A), Product (B), Sales (C)
            cells["A1"].PutValue("Region");
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
                cells[i + 1, 0].PutValue(data[i, 0]); // Region
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Sales
            }

            // Determine the last row and column that contain data
            int lastDataRow = cells.MaxDataRow;          // zero‑based index of the last used row
            int lastDataColumn = cells.MaxDataColumn;    // zero‑based index of the last used column

            // Define the cell area covering the entire data range (including header)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastDataRow,
                EndColumn = lastDataColumn
            };

            // Apply subtotal:
            // - Group by the first column (Region) -> groupBy = 0
            // - Use SUM function for subtotals
            // - Add subtotal for the third column (Sales) -> totalList = new int[] { 2 }
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 });

            // Save the workbook
            workbook.Save("SubtotalFullRange.xlsx");
        }
    }
}