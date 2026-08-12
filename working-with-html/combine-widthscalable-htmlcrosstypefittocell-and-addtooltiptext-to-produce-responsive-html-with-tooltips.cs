// Title: Responsive HTML Export with Tooltips (WidthScalable & FitToCell) – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert long and short text, set narrow column widths, and save it as responsive HTML using Aspose.Cells. The HtmlSaveOptions are configured with WidthScalable=true, HtmlCrossStringType=FitToCell, and AddTooltipText=true, producing a page that scales on any device, clips overflow text, and shows the full value on hover.
// Keywords: Aspose.Cells HTML export | WidthScalable true | HtmlCrossStringType FitToCell | AddTooltipText tooltip | responsive HTML C# | Excel to HTML Aspose | .NET web reporting | mobile‑friendly spreadsheet view | global Aspose.Cells example
// Common Searches: Aspose.Cells generate responsive HTML with tooltips | FitToCell clipping and tooltip in Aspose.Cells HTML export | How to enable WidthScalable in Aspose.Cells C# | Export Excel to HTML that adapts to screen size | Tooltip for truncated cell text Aspose.Cells
// Developer Intent: Export a workbook to HTML that automatically scales column widths, truncates overflow text to the cell boundary, and reveals the hidden content via hover tooltips.
// Use Cases: Publish a data‑rich report on mobile devices where long descriptions are shortened but still accessible. | Create HTML invoices with narrow columns; amounts that exceed column width are clipped yet viewable on hover. | Build an interactive dashboard that adapts to various screen sizes while preserving full text via tooltips.
// AI Prompts: Generate C# code to save an Aspose.Cells workbook as responsive HTML with WidthScalable, FitToCell, and AddTooltipText enabled. | Explain how HtmlCrossStringType.FitToCell affects text rendering in exported HTML and how tooltips are added. | Provide troubleshooting steps when tooltips do not appear after setting AddTooltipText = true in Aspose.Cells HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert long and short text, set narrow column widths, and save it as responsive HTML using Aspose.Cells. The HtmlSaveOptions are configured with WidthScalable=true, HtmlCrossStringType=FitToCell, and AddTooltipText=true, producing a page that scales on any device, clips overflow text, and shows the full value on hover.
    public class ResponsiveHtmlWithTooltipsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that may exceed cell width
            sheet.Cells["A1"].PutValue("This is a very long text that will not fit in the default column width and should show a tooltip.");
            sheet.Cells["B1"].PutValue("Short text");
            sheet.Cells["A2"].PutValue("Another long piece of text that demonstrates the FitToCell behavior.");

            // Set narrow column widths to force overflow
            sheet.Cells.SetColumnWidth(0, 10); // Column A
            sheet.Cells.SetColumnWidth(1, 8);  // Column B

            // Configure HTML save options:
            // - WidthScalable = true   => column widths are exported as scalable units (responsive)
            // - HtmlCrossStringType = FitToCell => text is clipped to cell width
            // - AddTooltipText = true => tooltip appears when text is clipped
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                WidthScalable = true,
                HtmlCrossStringType = HtmlCrossType.FitToCell,
                AddTooltipText = true
            };

            // Save the workbook as an HTML file with the configured options
            string outputPath = "ResponsiveWithTooltips.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
