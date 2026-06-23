using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ----- Populate sample data -----
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

        // ----- Define the range that will receive subtotals (A1:C6) -----
        CellArea range = CellArea.CreateCellArea("A1", "C6");

        // ----- Add subtotals:
        // Group by column 0 (Region), use SUM function on column 2 (Sales)
        // Replace existing subtotals, add page breaks, place summary below data
        cells.Subtotal(
            range,
            0,                                 // groupBy column index
            ConsolidationFunction.Sum,         // subtotal function
            new int[] { 2 },                   // columns to subtotal
            true,                              // replace existing subtotals
            true,                              // add page breaks between groups
            true                               // place summary below data
        );

        // ----- Freeze panes so that the header and subtotal rows stay visible -----
        // After applying subtotals, the first two rows (header + first subtotal) are at the top.
        // Freeze rows up to row 3 (i.e., rows 1‑2) and columns A‑C.
        worksheet.FreezePanes("A3", 2, 3);

        // ----- Save the workbook -----
        workbook.Save("SubtotalAndFreeze.xlsx");
    }
}