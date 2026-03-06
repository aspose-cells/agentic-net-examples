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

            // Populate sample data (header + rows)
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

            for (int r = 0; r < data.GetLength(0); r++)
                for (int c = 0; c < data.GetLength(1); c++)
                    cells[r + 1, c].PutValue(data[r, c]);

            // Define the range that contains the data (A1:C6)
            CellArea area = CellArea.CreateCellArea("A1", "C6");

            // Add subtotals:
            // - Group by the first column (Region) -> index 0
            // - Use SUM function
            // - Apply subtotal to the third column (Sales) -> index 2
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

            // Save the workbook in XLSX format
            workbook.Save("SubtotalDemo.xlsx");
        }
    }
}