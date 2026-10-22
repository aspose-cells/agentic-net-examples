// Title: Generate HTML with scalable column widths and cell comment tooltips using Aspose.Cells for .NET
// AI Prompts: Create an HTML file from a Workbook where column widths adjust automatically and cell comments appear as hover tooltips by configuring HtmlSaveOptions.WidthScalable and HtmlSaveOptions.AddTooltipText. | Export only the active worksheet to HTML with responsive columns and tooltip text using Aspose.Cells HtmlSaveOptions in C#.
// Common Searches: Aspose.Cells how to make column widths responsive in HTML export | C# export worksheet to HTML with cell comments as tooltips | Enable WidthScalable and AddTooltipText in Aspose.Cells HtmlSaveOptions example | Save only the active sheet to HTML using Aspose.Cells .NET | HTML export with scalable columns and hover tooltips Aspose.Cells tutorial
// Tags: HtmlSaveOptions.WidthScalable property | HtmlSaveOptions.AddTooltipText option | export active worksheet to HTML Aspose.Cells | scalable column widths HTML Aspose.Cells | cell comment tooltips HTML Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, fills it with product data, sets column widths, configures HtmlSaveOptions to enable WidthScalable and AddTooltipText, limits the export to the active worksheet, and saves the result as an HTML file with responsive columns and hover tooltips.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["C2"].PutValue(0.5);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(85);
        sheet.Cells["C3"].PutValue(0.3);

        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(60);
        sheet.Cells["C4"].PutValue(1.2);

        // Optionally, set column widths to demonstrate scalability
        sheet.Cells.SetColumnWidth(0, 20); // Column A
        sheet.Cells.SetColumnWidth(1, 15); // Column B
        sheet.Cells.SetColumnWidth(2, 15); // Column C

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Enable scalable column widths when the HTML is displayed
            WidthScalable = true,

            // Add tooltip text (cell comments) to HTML cells
            AddTooltipText = true,

            // Optional: Export only the first worksheet
            ExportActiveWorksheetOnly = true
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ScalableColumnsWithTooltips.html", saveOptions);
    }
}
