// Title: Calculate average subtotal for column F and place summary rows at the top with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills column F with numeric values, defines the range A1:F6, and calls Cells.Subtotal to compute the average of column F grouped by the first column. By setting the summaryBelowData parameter to false, the subtotal rows are inserted above each group, giving a top‑summary view.
// Keywords: Aspose.Cells subtotal average | C# Excel subtotal top rows | Aspose.Cells Cells.Subtotal | summaryBelowData false | .NET Excel grouping | average function Aspose.Cells | place subtotal rows above data | CellArea CreateCellArea example | ConsolidationFunction.Average
// Common Searches: Aspose.Cells add average subtotal at top of column | C# place subtotal summary rows above data in Excel | how to use Cells.Subtotal summaryBelowData false | group by first column and show average before details Aspose.Cells | Excel subtotal top rows C# Aspose
// Developer Intent: Insert an average subtotal for column F and have the summary rows appear above the data rows.
// Use Cases: Financial statements where the average amount is shown before the detailed transaction list. | Sales dashboards that display the average sales per region at the top of each region group. | Inventory reports that provide a quick average quantity summary before listing individual items.
// AI Prompts: Show C# code using Aspose.Cells to add an average subtotal for column F with the summary rows positioned at the top. | Explain how the summaryBelowData flag in Cells.Subtotal controls whether subtotal rows appear above or below the data. | Give a step‑by‑step guide to group rows by the first column and calculate the average of column F using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example creates a workbook, fills column F with numeric values, defines the range A1:F6, and calls Cells.Subtotal to compute the average of column F grouped by the first column. By setting the summaryBelowData parameter to false, the subtotal rows are inserted above each group, giving a top‑summary view.
class SubtotalAverageTopExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column F (zero‑based column index 5)
            // Header
            cells["F1"].PutValue("Amount");
            // Data rows (F2:F6)
            cells["F2"].PutValue(120);
            cells["F3"].PutValue(150);
            cells["F4"].PutValue(200);
            cells["F5"].PutValue(180);
            cells["F6"].PutValue(170);

            // Define the range that contains the data (including header)
            // From A1 (row 0, column 0) to F6 (row 5, column 5)
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 5);

            // Add subtotals:
            // - Group by the first column (index 0)
            // - Use Average function
            // - Apply the subtotal to column F (index 5)
            // - Do not replace existing subtotals, no page breaks, and place summary below data (summaryBelowData = true)
            cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Average,    // average function
                new int[] { 5 },                  // subtotal column (F)
                false,                            // replace existing subtotals
                false,                            // add page breaks between groups
                true);                            // summaryBelowData = true => summary rows at the bottom

            // Save the workbook
            workbook.Save("Subtotal_Average_Top.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
