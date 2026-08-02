// Title: Aspose.Cells HTML Export – Prevent tooltip on short text when AddTooltipText is true
// Description: Demonstrates how to export a workbook to HTML with HtmlSaveOptions.AddTooltipText enabled so that only cells whose content is truncated (e.g., long text in A2) receive a tooltip attribute, while cells that fit the column width (e.g., short text in A1) do not.
// Keywords: Aspose.Cells HTML tooltip | AddTooltipText short text | HTML export truncated cell tooltip | Aspose.Cells column width tooltip | C# Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells tooltip only for overflow cells | HTML export AddTooltipText short text no tooltip | prevent title attribute on short cells Aspose | how to disable tooltip for fitting text in Aspose.Cells HTML | Aspose.Cells generate tooltips only when text is cut off
// Developer Intent: Create HTML output where tooltip attributes are added exclusively to cells whose displayed text is clipped by column width.
// Use Cases: Web dashboards that show tooltips only for truncated spreadsheet values, reducing visual noise. | Automated validation that short cells lack the title attribute while long cells include it after HTML conversion. | Generating printable HTML reports where unnecessary hover pop‑ups are avoided.
// AI Prompts: Explain how HtmlSaveOptions.AddTooltipText decides which cells get a tooltip in Aspose.Cells for .NET. | Provide a C# example that checks the generated HTML to confirm that short cells have no title attribute. | Show how to adjust column width and AddTooltipText to ensure tooltips appear only for overflow text.

using System;
using Aspose.Cells;

namespace AsposeCellsTooltipDemo
{
    // Demonstrates how to export a workbook to HTML with HtmlSaveOptions.AddTooltipText enabled so that only cells whose content is truncated (e.g., long text in A2) receive a tooltip attribute, while cells that fit the column width (e.g., short text in A1) do not.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Short text that fits within the column width
            cells["A1"].PutValue("Short");

            // Long text that exceeds the column width and will need a tooltip
            cells["A2"].PutValue("This is a very long text that will not fit in the column and should display a tooltip");

            // Set a narrow column width so that the long text is truncated
            cells.SetColumnWidth(0, 10); // Width in characters

            // Configure HTML save options to add tooltip text when data is truncated
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.AddTooltipText = true; // Enable tooltip generation

            // Save the workbook as HTML
            string outputPath = "TooltipDemo.html";
            workbook.Save(outputPath, saveOptions);

            // At this point, cell A1 (short text) will not have a tooltip attribute,
            // while cell A2 (long text) will have a tooltip because its content cannot be fully displayed.
            Console.WriteLine("Workbook saved to " + outputPath);
        }
    }
}
