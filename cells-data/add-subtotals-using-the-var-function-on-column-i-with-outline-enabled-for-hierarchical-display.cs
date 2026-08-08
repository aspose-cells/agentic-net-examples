// Title: C# – Add variance subtotals with outline hierarchy using Aspose.Cells
// Description: Shows how to create a workbook, fill Category (H) and Value (I) columns, define a CellArea, and call Cells.Subtotal with ConsolidationFunction.Var to calculate variance subtotals grouped by Category. The example enables outline view, places summary rows below detail rows, inserts page breaks, and saves the file as SubtotalVarOutlineDemo.xlsx.
// Keywords: Aspose.Cells C# subtotal | variance subtotal Aspose | ConsolidationFunction.Var example | .NET Excel outline view | group by column subtotal | summary row below outline | page break between groups | CellArea Subtotal method
// Common Searches: Aspose.Cells subtotal Var function C# | enable outline view for subtotals Aspose.Cells | how to add variance subtotals in Excel with Aspose | group data by category and calculate variance using Aspose.Cells | C# code for hierarchical subtotals with page breaks
// Developer Intent: Generate an Excel workbook that groups rows by Category, computes variance subtotals for the Value column, and presents the results in an expandable outline.
// Use Cases: Financial statements that need variance analysis per department with collapsible sections. | Inventory reports showing variance of stock levels for each product category. | Sales dashboards that automatically calculate and hide variance subtotals for regional groups.
// AI Prompts: Write C# code with Aspose.Cells to add variance subtotals on column I and enable outline view with summary rows below the data. | Explain how to modify the Subtotal method parameters to place summary rows above the detail rows while still using the Var function. | Provide a step‑by‑step guide to insert page breaks between groups when creating variance subtotals with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalVarOutlineDemo
{
    // Shows how to create a workbook, fill Category (H) and Value (I) columns, define a CellArea, and call Cells.Subtotal with ConsolidationFunction.Var to calculate variance subtotals grouped by Category. The example enables outline view, places summary rows below detail rows, inserts page breaks, and saves the file as SubtotalVarOutlineDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in columns H (index 7) and I (index 8)
            // Header row
            cells[0, 7].PutValue("Category");   // Column H
            cells[0, 8].PutValue("Value");      // Column I

            // Sample data rows
            object[,] data = new object[,]
            {
                { "A", 10 },
                { "A", 20 },
                { "B", 30 },
                { "B", 40 },
                { "C", 50 },
                { "C", 60 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                cells[r + 1, 7].PutValue(data[r, 0]); // Category
                cells[r + 1, 8].PutValue(data[r, 1]); // Value
            }

            // Define the cell area that contains the data (including headers)
            // StartRow = 0, StartColumn = 7 (H), EndRow = data rows count, EndColumn = 8 (I)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 7,
                EndRow = data.GetLength(0),
                EndColumn = 8
            };

            // Add subtotals:
            // - Group by the first column in the area (Category) => groupBy = 0
            // - Use Var function for variance calculation
            // - Apply subtotal to the second column in the area (Value) => totalList = {1}
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                area,
                0,
                ConsolidationFunction.Var,
                new int[] { 1 },
                true,   // replace existing subtotals
                true,   // add page breaks between groups
                true    // summary below data (hierarchical outline)
            );

            // Enable outline view and ensure summary rows appear below detail rows
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalVarOutlineDemo.xlsx");
        }
    }
}
