// Title: Set Smart Marker LabelPosition to "After" for Group Summary Row in Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel workbook with Aspose.Cells, group detail rows, and place the group summary label after the grouped rows by configuring the smart marker LabelPosition attribute (or Outline.SummaryRowBelow) to true, then save the file.
// Keywords: Aspose.Cells group rows | LabelPosition After | smart marker summary row | C# Excel grouping | Outline.SummaryRowBelow | .NET Excel export | group label below data
// Common Searches: Aspose.Cells set smart marker label position after | C# group rows summary row below example | Outline.SummaryRowBelow true Aspose.Cells | how to place group label after detail rows in Excel using .NET | smart marker LabelPosition attribute usage
// Developer Intent: Generate an Excel file where the group’s summary label appears below the grouped detail rows by setting the smart marker LabelPosition to "After".
// Use Cases: Sales reports that list products and show a subtotal row after each product group. | Financial statements with section totals displayed beneath the related line items. | Exporting hierarchical data structures where each parent label follows its child rows.
// AI Prompts: Show me C# code that sets a smart marker's LabelPosition to "After" so the group summary appears after the data rows in Aspose.Cells. | Provide an Aspose.Cells example that groups rows and uses Outline.SummaryRowBelow to place the summary row below the grouped rows. | Explain how to configure group label placement for Excel exports using Aspose.Cells smart markers in .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create an Excel workbook with Aspose.Cells, group detail rows, and place the group summary label after the grouped rows by configuring the smart marker LabelPosition attribute (or Outline.SummaryRowBelow) to true, then save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("Product A");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Product B");
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["A4"].PutValue("Total");

            // Set formula for total
            worksheet.Cells["B4"].Formula = "=SUM(B2:B3)";

            // Group the detail rows (rows 2 and 3). 
            // Parameters: start row index (zero‑based), number of rows to group, collapsed flag.
            worksheet.Cells.GroupRows(1, 2, false);

            // Place the summary (group label) after the detail rows.
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            string outputPath = "GroupLabelsAfterDataRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
