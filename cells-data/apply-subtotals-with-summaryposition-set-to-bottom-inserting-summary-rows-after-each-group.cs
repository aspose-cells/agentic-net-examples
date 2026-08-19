// Title: Create Bottom‑Positioned Subtotal Rows per Group with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells' Cells.Subtotal method to group data by the first column, sum the Amount column, and insert subtotal rows after each category. The Outline.SummaryRowBelow property is set to true so the summary rows appear below the detail rows, and the workbook is saved as SubtotalBottomDemo.xlsx.
// Keywords: Aspose.Cells | C# subtotal | summary row below | outline summary row | group by column | Excel subtotal | Cells.Subtotal | bottom subtotal | Aspose.Cells example | Excel automation .NET
// Common Searches: Aspose.Cells add subtotal rows at the bottom of each group | C# subtotal summary row below example | Cells.Subtotal method outline summary row below | group by column subtotal Aspose.Cells .NET | insert subtotal rows after each category using Aspose.Cells
// Developer Intent: Insert subtotal rows after each grouped category with the summary placed below the detail rows using Aspose.Cells in C#.
// Use Cases: Financial statements that show a total for each expense category directly beneath its line items. | Inventory reports where each supplier's items are followed by a subtotal of quantities or values. | Sales dashboards that display regional totals after the last transaction of each region.
// AI Prompts: Generate C# code with Aspose.Cells to apply subtotals and place the summary row below each group for a given range. | Explain each parameter of the Cells.Subtotal method and how Outline.SummaryRowBelow controls the position of subtotal rows. | Show how to extend the example to subtotal multiple columns and customize the formatting of the bottom summary rows.

using Aspose.Cells;
using System;

// Demonstrates how to use Aspose.Cells' Cells.Subtotal method to group data by the first column, sum the Amount column, and insert subtotal rows after each category. The Outline.SummaryRowBelow property is set to true so the summary rows appear below the detail rows, and the workbook is saved as SubtotalBottomDemo.xlsx.
class SubtotalBottomDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Item");
        cells["C1"].PutValue("Amount");

        object[,] data = new object[,]
        {
            { "A", "Item1", 100 },
            { "A", "Item2", 150 },
            { "B", "Item3", 200 },
            { "B", "Item4", 250 },
            { "A", "Item5", 120 }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]);
            cells[i + 1, 1].PutValue(data[i, 1]);
            cells[i + 1, 2].PutValue(data[i, 2]);
        }

        // Define the range that contains the data (A1:C6)
        CellArea area = CellArea.CreateCellArea("A1", "C6");

        // Apply subtotals:
        // - Group by the first column (Category) -> groupBy = 0
        // - Use SUM function for subtotals
        // - Subtotal the third column (Amount) -> totalList = {2}
        // - Do not replace existing subtotals, no page breaks, place summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, false, false, true);

        // Ensure the outline places summary rows below the detail rows
        worksheet.Outline.SummaryRowBelow = true;

        // Save the workbook
        workbook.Save("SubtotalBottomDemo.xlsx");
    }
}
