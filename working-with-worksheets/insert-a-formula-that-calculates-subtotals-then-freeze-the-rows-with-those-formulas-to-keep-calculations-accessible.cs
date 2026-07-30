// Title: Add Subtotal Formulas and Freeze Panes with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates sales data, inserts subtotal rows that sum the Sales column grouped by Region, freezes the rows containing the first subtotal to keep calculations visible, and saves the file as SubtotalFreezeDemo.xlsx.
// Keywords: Aspose.Cells subtotal C# | freeze panes Aspose.Cells | Excel subtotal formula .NET | grouped subtotals Aspose.Cells | freeze rows after subtotal
// Common Searches: Aspose.Cells add subtotal rows C# | freeze panes after subtotal Aspose.Cells | how to use Subtotal method Aspose.Cells .NET | keep subtotal visible Excel freeze panes C#
// Developer Intent: Insert grouped subtotal formulas into a worksheet and freeze the rows that contain those subtotals so they remain in view while scrolling.
// Use Cases: Generate a regional sales report that automatically calculates totals and keeps them visible during navigation. | Create an invoice workbook where category subtotals are frozen at the top of each section for quick reference. | Build a financial statement with grouped subtotals and frozen panes to improve readability of large data sets.
// AI Prompts: Show how to add multiple subtotal rows for different columns using Aspose.Cells in C#. | Explain how to calculate the number of rows to freeze after inserting subtotal rows. | Provide code to format subtotal rows and then freeze the panes in an Excel file with Aspose.Cells.

using Aspose.Cells;
using System;

// Creates a workbook, populates sales data, inserts subtotal rows that sum the Sales column grouped by Region, freezes the rows containing the first subtotal to keep calculations visible, and saves the file as SubtotalFreezeDemo.xlsx.
class SubtotalFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data
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
            cells[i + 1, 0].PutValue(data[i, 0]);
            cells[i + 1, 1].PutValue(data[i, 1]);
            cells[i + 1, 2].PutValue(data[i, 2]);
        }

        // Define the range that contains the data (including header)
        CellArea area = CellArea.CreateCellArea("A1", "C5");

        // Add subtotals:
        // - Group by column 0 (Region)
        // - Use SUM function on column 2 (Sales)
        // - Place summary rows below the data (summaryBelowData = true)
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, false, false, true);

        // Freeze the rows that contain the subtotal formulas.
        // In this example the first subtotal row appears after the first group,
        // so we freeze the top 4 rows (header + first group) to keep the subtotal visible.
        worksheet.FreezePanes(4, 0, 4, 0);

        // Save the workbook
        workbook.Save("SubtotalFreezeDemo.xlsx");
    }
}
