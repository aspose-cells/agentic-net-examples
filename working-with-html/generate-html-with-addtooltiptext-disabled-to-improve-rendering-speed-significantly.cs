// Title: Aspose.Cells C# – Export Excel to HTML without Tooltips for Faster Rendering
// Description: Demonstrates how to save a workbook as HTML using Aspose.Cells with HtmlSaveOptions.AddTooltipText set to false, eliminating tooltip markup and significantly speeding up page load.
// Keywords: Aspose.Cells HTML export C# | AddTooltipText false | disable tooltips Aspose.Cells | HTML rendering performance | Excel to HTML without tooltips
// Common Searches: Aspose.Cells turn off tooltips when saving to HTML | HtmlSaveOptions AddTooltipText performance impact | C# export Excel to HTML faster Aspose.Cells | How to remove tooltip text from HTML output Aspose.Cells
// Developer Intent: Export an Excel workbook to HTML while suppressing tooltip generation to improve export speed.
// Use Cases: Create lightweight HTML previews of large spreadsheets where overflow tooltips are unnecessary. | Generate fast‑loading HTML reports in web apps by omitting tooltip markup. | Produce mobile‑friendly HTML exports of data‑heavy workbooks with reduced payload.
// AI Prompts: Show C# code that saves an Aspose.Cells workbook to HTML with AddTooltipText disabled. | Explain how HtmlSaveOptions.AddTooltipText affects HTML size and rendering time. | Provide a step‑by‑step guide to improve Aspose.Cells HTML export performance by turning off tooltips.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlTooltipDemo
{
    // Demonstrates how to save a workbook as HTML using Aspose.Cells with HtmlSaveOptions.AddTooltipText set to false, eliminating tooltip markup and significantly speeding up page load.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data that would normally require a tooltip if it overflows
            sheet.Cells["A1"].PutValue("This is a very long text that would normally need a tooltip when displayed in HTML.");

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 10);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Disable tooltip generation to improve rendering speed
            htmlOptions.AddTooltipText = false;

            // Save the workbook as HTML with the specified options
            string outputPath = "output_without_tooltip.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with AddTooltipText disabled.");
        }
    }
}
