// Title: Add Subtotal Rows Above Groups and Freeze Header in Aspose.Cells for .NET
// Description: Creates a workbook, fills it with region, product, and sales data, inserts SUM subtotals above each region group using Worksheet.Cells.Subtotal, freezes the header row and the first subtotal row with FreezePanes, and saves the file as XLSX.
// Keywords: Aspose.Cells subtotal | C# FreezePanes | group by column Aspose.Cells | insert subtotal rows above data | freeze header row Excel .NET | Worksheet.Cells.Subtotal example | Aspose.Cells sales report
// Common Searches: Aspose.Cells add subtotal rows above each group | freeze header and subtotal rows with Aspose.Cells C# | Worksheet.Cells.Subtotal usage .NET | how to freeze panes in Aspose.Cells | C# example subtotal and freeze panes
// Developer Intent: Insert SUM subtotals above grouped rows and lock the top rows so they stay visible while scrolling.
// Use Cases: Generate a sales summary that shows region‑level totals before the detailed rows and keeps those totals in view. | Create a financial worksheet with grouped subtotals placed above each section and a frozen header for quick reference. | Prepare large Excel exports where grouped subtotals are needed and the first rows must remain static during navigation.
// AI Prompts: Write C# code with Aspose.Cells to add subtotal rows above each group and freeze the first two rows. | Explain each parameter of Worksheet.Cells.Subtotal and how to control grouping, function type, and row placement. | Show how to use FreezePanes in Aspose.Cells to lock multiple rows and columns at once.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalAndFreezeDemo
{
    // Creates a workbook, fills it with region, product, and sales data, inserts SUM subtotals above each region group using Worksheet.Cells.Subtotal, freezes the header row and the first subtotal row with FreezePanes, and saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (Region, Product, Sales)
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                {"North", "Widget", 5000},
                {"North", "Gadget", 3000},
                {"South", "Widget", 6000},
                {"South", "Gadget", 4000},
                {"West",  "Widget", 4500}
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Region
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Sales
            }

            // Define the range that contains the data (A1:C6)
            CellArea area = CellArea.CreateCellArea("A1", "C6");

            // Add subtotals:
            // - Group by the first column (Region) -> index 0
            // - Use SUM function on the Sales column -> index 2
            // - Place subtotal rows above each group (summaryBelowData = false)
            // - Replace existing subtotals, no page breaks
            worksheet.Cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 2 },                   // columns to subtotal (Sales)
                true,                              // replace existing subtotals
                false,                             // no page breaks
                false);                            // place subtotal above data

            // After placing subtotals above each group, the first subtotal row is row 2 (zero‑based index 1)
            // Freeze the header row and the first subtotal row so they stay visible while scrolling
            // FreezePanes(rowIndex, columnIndex, freezedRows, freezedColumns)
            worksheet.FreezePanes(2, 0, 2, 0); // Freeze rows 0 and 1 (header + first subtotal)

            // Save the workbook
            workbook.Save("SubtotalAndFreezeDemo.xlsx");
        }
    }
}
