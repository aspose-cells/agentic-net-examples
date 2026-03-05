using Aspose.Cells;
using System;

class SubtotalDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row
        cells["A1"].PutValue("Region");
        cells["B1"].PutValue("Product");
        cells["C1"].PutValue("Sales");

        // Populate sample data
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

        // Define the cell area that contains the data (A1:C6)
        CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

        // Add subtotals:
        // - Group by column 0 (Region)
        // - Use SUM function
        // - Apply subtotal to column 2 (Sales)
        // - Replace existing subtotals, add page breaks, place summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

        // Save the workbook in XLSX format
        workbook.Save("SubtotalDemo.xlsx");
    }
}