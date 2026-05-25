using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalAfterSort
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header)
            // Column A: Category, Column B: Amount
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            object[,] data = new object[,]
            {
                { "B", 200 },
                { "A", 100 },
                { "C", 300 },
                { "B", 150 },
                { "A", 250 },
                { "C", 350 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                cells[r + 1, 0].PutValue(data[r, 0]); // Category
                cells[r + 1, 1].PutValue(data[r, 1]); // Amount
            }

            // Define the range that contains the data (including header)
            CellArea dataArea = CellArea.CreateCellArea("A1", $"B{data.GetLength(0) + 1}");

            // Sort the data by the first column (Category) in ascending order
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;               // First row is a header
            sorter.Key1 = 0;                         // Sort by column A (index 0)
            sorter.Order1 = SortOrder.Ascending;     // Ascending order
            sorter.Sort(cells, dataArea);            // Perform the sort

            // Apply subtotals:
            // - Group by the first column (Category) -> groupBy = 0
            // - Use SUM function on the second column (Amount) -> totalList = new int[] { 1 }
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                dataArea,
                0,                                 // Group by column A
                ConsolidationFunction.Sum,         // Sum function
                new int[] { 1 },                   // Subtotal on column B
                true,                              // Replace existing subtotals
                true,                              // Add page breaks between groups
                true                               // Summary below data
            );

            // Save the workbook
            workbook.Save("SubtotalAfterSort.xlsx");
        }
    }
}