using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalAfterSort
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Populate sample data (including header) in columns A‑C
            // Header
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            // Data rows (unsorted on purpose)
            object[,] data = new object[,]
            {
                { "West",  "Widget", 4500 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "North", "Widget", 5000 },
                { "South", "Gadget", 4000 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // 3. Define the range that contains the data (including header)
            // A1:C6  -> rows 0‑5, columns 0‑2
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

            // 4. Sort the range by the first column (Region) in ascending order
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                 // first row is a header
            sorter.Key1 = 0;                           // column A (zero‑based)
            sorter.Order1 = SortOrder.Ascending;       // ascending
            sorter.Sort(worksheet.Cells, area);        // perform the sort

            // 5. Apply subtotals after sorting
            // Group by column A (Region), sum the Sales column (index 2)
            // Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                area,
                0,                                 // groupBy column index (Region)
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 2 },                   // columns to subtotal (Sales)
                true,                              // replace existing subtotals
                true,                              // add page breaks between groups
                true                               // place summary below data
            );

            // 6. Save the workbook
            workbook.Save("SubtotalAfterSort.xlsx");
        }
    }
}