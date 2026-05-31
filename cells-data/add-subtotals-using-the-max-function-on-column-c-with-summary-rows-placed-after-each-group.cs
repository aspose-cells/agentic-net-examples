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

            // Populate sample data (Header + some rows)
            cells["A1"].PutValue("Category");   // Column A
            cells["B1"].PutValue("Item");       // Column B
            cells["C1"].PutValue("Amount");     // Column C (the column we will subtotal)

            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 },
                { "West",  "Gadget", 3500 }
            };

            // Fill the worksheet with the data starting from row 2 (zero‑based index 1)
            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the range that contains the data (including header)
            // A1:C7  => rows 0‑6, columns 0‑2
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0),   // header row + data rows - 1
                EndColumn = 2
            };

            // Add subtotals:
            // - Group by the first column (Category) => index 0
            // - Use Max function on column C (Amount) => index 2
            // - Replace existing subtotals: true
            // - No page breaks between groups: false
            // - Place summary rows below each group: true
            cells.Subtotal(
                area,
                0,                                 // groupBy column (Category)
                ConsolidationFunction.Max,         // Max function
                new int[] { 2 },                   // subtotal on column C (Amount)
                true,                              // replace existing subtotals
                false,                             // no page breaks
                true                               // summary rows below data
            );

            // Save the workbook
            workbook.Save("Subtotal_Max_ColumnC.xlsx");
        }
    }
}