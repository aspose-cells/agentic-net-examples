// Title: Create HTML from an Excel workbook with scalable column widths and hover tooltips using Aspose.Cells for .NET
// AI Prompts: Configure HtmlSaveOptions with WidthScalable = true and AddTooltipText = true, then save the workbook as an HTML file. | Populate a worksheet, set specific column widths, and export it to HTML while preserving column scaling and showing cell comments on hover. | Generate an HTML representation of a workbook that maintains responsive column sizes and displays tooltip text for each cell comment.
// Common Searches: Aspose.Cells how to enable column scaling in HTML export | Add tooltip for cell comments when saving Excel as HTML with Aspose.Cells | C# export Excel to HTML with responsive column widths using HtmlSaveOptions | Example of WidthScalable and AddTooltipText properties in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions WidthScalable | Aspose.Cells HtmlSaveOptions AddTooltipText | export Excel to HTML with responsive columns | cell comments as hover tooltips Aspose.Cells | C# set column width for HTML output

using System;
using Aspose.Cells;

// The program creates a workbook, fills it with sample data, adjusts column widths, enables WidthScalable and AddTooltipText in HtmlSaveOptions, and saves the result as an HTML file that features scalable columns and hover tooltips for cell comments.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.20);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.80);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(2.50);

        // Optionally set column widths (demonstrates scalability)
        sheet.Cells.SetColumnWidth(0, 20); // Column A
        sheet.Cells.SetColumnWidth(1, 15); // Column B

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Enable scalable column widths when the HTML is rendered
            WidthScalable = true,

            // Add tooltip text (cell comments) to HTML cells on hover
            AddTooltipText = true
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ScalableColumnsWithTooltips.html", htmlOptions);
    }
}
