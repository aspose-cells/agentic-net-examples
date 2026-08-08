// Title: Create Max Subtotals by Category with Post‑Group Summary Rows using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate an Excel workbook, populate it with sample data, define a cell range, and use the Cells.Subtotal method to group rows by the first column, calculate the maximum value of column C for each group, and insert subtotal rows below each group before saving the file.
// Keywords: Aspose.Cells | C# | .NET | Subtotal method | Max function | group by column | summary rows | Excel automation | CellArea | ConsolidationFunction.Max | worksheet subtotal
// Common Searches: Aspose.Cells add max subtotal per category | C# Subtotal method summary rows after groups | How to group rows and calculate max in Excel using Aspose.Cells | Create subtotal rows below each group with Aspose.Cells .NET | Excel max subtotal example Aspose.Cells
// Developer Intent: Insert Max‑based subtotal rows for each category, positioned after the grouped rows.
// Use Cases: Sales dashboard that shows the highest sale amount per product line with a subtotal row beneath each line. | Inventory report displaying the maximum stock level per warehouse section, followed by a summary row. | Financial statement that lists the peak expense for each department, with a subtotal row placed directly under each department group.
// AI Prompts: Generate C# code with Aspose.Cells to add subtotal rows that compute the maximum of column D grouped by column B, placing the summary rows after each group. | Explain the purpose of each parameter in the Cells.Subtotal method, including grouping column index, aggregation function, target columns, and summary row placement. | Show how to apply a custom style (font, background color, borders) to the subtotal rows created by the Subtotal method in Aspose.Cells.

using Aspose.Cells;
using System;

// Demonstrates how to generate an Excel workbook, populate it with sample data, define a cell range, and use the Cells.Subtotal method to group rows by the first column, calculate the maximum value of column C for each group, and insert subtotal rows below each group before saving the file.
class SubtotalMaxDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add header row
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Item");
        cells["C1"].PutValue("Value");

        // Sample data (Category, Item, Value)
        object[,] data = new object[,]
        {
            { "A", "Item1", 10 },
            { "A", "Item2", 20 },
            { "B", "Item3", 15 },
            { "B", "Item4", 25 },
            { "A", "Item5", 30 }
        };

        // Populate the worksheet with the sample data
        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]); // Column A
            cells[i + 1, 1].PutValue(data[i, 1]); // Column B
            cells[i + 1, 2].PutValue(data[i, 2]); // Column C
        }

        // Define the cell area that includes the header and all data rows (A1:C6)
        CellArea area = CellArea.CreateCellArea(0, 0, data.GetLength(0), 2);

        // Add subtotals:
        // - Group by the first column (Category) -> groupBy = 0
        // - Use Max function on column C (zero‑based index 2)
        // - Replace existing subtotals if any, no page breaks, summary rows placed below each group
        cells.Subtotal(area, 0, ConsolidationFunction.Max, new int[] { 2 }, true, false, true);

        // Save the workbook
        workbook.Save("SubtotalMaxDemo.xlsx");
    }
}
