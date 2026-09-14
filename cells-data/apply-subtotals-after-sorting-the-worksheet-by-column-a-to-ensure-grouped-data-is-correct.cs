// Title: Sorting an Excel sheet by column A and inserting SUM subtotals for Sales grouped by Category with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to sort a worksheet range by the first column (including headers) and then adds SUM subtotals for a specified column, grouping by the sorted column. | Show how to apply page breaks and place subtotal rows below each group after sorting data with Aspose.Cells in a .NET application. | Create an Aspose.Cells example that populates sample data, sorts it by Category, adds grouped subtotals on the Sales column, and saves the workbook.
// Common Searches: Aspose.Cells C# sort data range and then add grouped subtotals | how to apply SUM subtotals after sorting worksheet with Aspose.Cells | C# example for inserting page breaks between subtotal groups using Aspose.Cells | subtotal function on specific column after sorting with headers Aspose.Cells | group data by column and calculate totals in Excel via Aspose.Cells .NET
// Tags: sort worksheet range with DataSorter Aspose.Cells | apply Subtotal method after sorting Aspose.Cells | grouped SUM subtotals using ConsolidationFunction.Sum | insert page breaks between subtotal groups Aspose.Cells | place subtotal rows below each group Aspose.Cells | Excel data sorting and subtotaling in C# Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalAfterSort
{
    // The example creates a workbook, fills it with Category, Product, and Sales data, sorts the range by the Category column (A) while recognizing the header row, then adds SUM subtotals for the Sales column grouped by Category, inserts page breaks between groups, places summary rows below each group, and saves the file as SubtotalAfterSort.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header) in columns A, B, C
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            // Data rows
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

            // -------------------------------------------------
            // 1. Sort the data range by the first column (A)
            // -------------------------------------------------
            // Define the area to sort (including header)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0),   // last data row index (header + data)
                EndColumn = 2                // up to column C
            };

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;                 // first row is a header
            sorter.Key1 = 0;                          // sort by column A (zero‑based)
            sorter.Order1 = SortOrder.Ascending;      // ascending order

            // Perform the sort
            sorter.Sort(cells, sortArea);

            // -------------------------------------------------
            // 2. Apply subtotals on the sorted range
            // -------------------------------------------------
            // Define the area for subtotals (same as sort area)
            CellArea subtotalArea = sortArea;

            // Apply subtotal:
            // - Group by column 0 (Category)
            // - Use SUM function
            // - Subtotal the Sales column (index 2)
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                subtotalArea,
                0,                                 // group by first column
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 2 },                   // columns to subtotal (Sales)
                true,                              // replace existing subtotals
                true,                              // insert page breaks between groups
                true                               // place summary rows below each group
            );

            // -------------------------------------------------
            // 3. Save the workbook
            // -------------------------------------------------
            workbook.Save("SubtotalAfterSort.xlsx");
        }
    }
}
