// Title: C# – Sort Worksheet by Column A and Add Subtotals with Aspose.Cells
// Description: This example creates a workbook, fills it with a header and sample sales data, sorts the range A1:C6 by the Category column using Aspose.Cells' DataSorter, then applies Cells.Subtotal to group rows by Category, sum the Sales column, replace any existing subtotals, insert page breaks between groups, and place the summary rows below the details before saving the file.
// Keywords: Aspose.Cells C# sort | Aspose.Cells subtotal | DataSorter | Cells.Subtotal | group by column | insert page breaks | summary rows below data | Excel automation C# | subtotal after sort
// Common Searches: Aspose.Cells sort and subtotal example | C# add subtotals after sorting Excel data | How to use DataSorter with Subtotal in Aspose.Cells | Insert page breaks between subtotal groups C# | Group sales by category using Aspose.Cells
// Developer Intent: The developer wants to order rows by the Category field and then generate subtotal rows that total sales for each category, optionally adding page breaks and placing the totals beneath each group.
// Use Cases: Produce a sales report where categories are alphabetically ordered and each category shows a total sales row underneath its items. | Create an invoice summary that groups line items by product type after sorting, inserting page breaks for printable sections. | Generate a quick pivot‑like view without a pivot table by sorting data and applying the Subtotal method to calculate grouped totals.
// AI Prompts: Show how to subtotal multiple columns (e.g., Sales and Quantity) after sorting with Aspose.Cells. | Provide code that adds a grand total row after all category subtotals in a C# workbook. | Explain how to sort by several columns before applying the Subtotal method in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalAfterSort
{
    // This example creates a workbook, fills it with a header and sample sales data, sorts the range A1:C6 by the Category column using Aspose.Cells' DataSorter, then applies Cells.Subtotal to group rows by Category, sum the Sales column, replace any existing subtotals, insert page breaks between groups, and place the summary rows below the details before saving the file.
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
            // Step 1: Sort the data by the first column (Category)
            // -------------------------------------------------
            // Configure the DataSorter to sort by column 0 (Category) in ascending order
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true;               // First row contains headers
            sorter.Key1 = 0;                         // Sort by column A (zero‑based index)
            sorter.Order1 = SortOrder.Ascending;     // Ascending order

            // Define the area to sort (including header row)
            CellArea sortArea = CellArea.CreateCellArea("A1", "C6"); // rows 0‑5, cols 0‑2
            sorter.Sort(cells, sortArea);

            // -------------------------------------------------
            // Step 2: Apply subtotals on the sorted data
            // -------------------------------------------------
            // Define the range for subtotals (same as sort area)
            CellArea subtotalArea = sortArea;

            // Group by the first column (Category), sum the Sales column (index 2)
            // Replace existing subtotals, add page breaks between groups, place summary below data
            cells.Subtotal(
                subtotalArea,
                0,                                 // groupBy column index (Category)
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 2 },                   // columns to subtotal (Sales)
                true,                              // replace existing subtotals
                true,                              // insert page breaks between groups
                true                               // place summary rows below the detail rows
            );

            // -------------------------------------------------
            // Step 3: Save the workbook
            // -------------------------------------------------
            workbook.Save("SubtotalAfterSort.xlsx");
        }
    }
}
