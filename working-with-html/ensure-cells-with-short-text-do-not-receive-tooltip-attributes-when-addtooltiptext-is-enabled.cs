// Title: Aspose.Cells C# – Export to HTML without tooltips on short‑text cells (AddTooltipText enabled)
// Description: Shows how to save a workbook to HTML with HtmlSaveOptions.AddTooltipText turned on, while ensuring that only cells whose text overflows the column width receive a tooltip attribute, leaving fitting cells tooltip‑free.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | AddTooltipText | HTML export | tooltip overflow | short text cells | column width | prevent tooltip | cell overflow detection
// Common Searches: Aspose.Cells prevent tooltip on short text | AddTooltipText only for overflowed cells | HTML export tooltip overflow Aspose | C# Aspose.Cells hide tooltip for fitting cells | How to disable tooltip for short cells Aspose
// Developer Intent: Generate HTML with tooltips only for cells whose content exceeds the column width while keeping AddTooltipText enabled.
// Use Cases: Create web‑ready spreadsheets where tooltips appear only for truncated values. | Produce clean HTML reports that avoid unnecessary title attributes on short entries. | Implement conditional tooltip suppression for specific rows or columns in automated export pipelines. | Improve accessibility by limiting tooltip clutter to meaningful overflowed data.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to HTML with AddTooltipText enabled, but automatically omit the title attribute for cells that fit within the column width. | Explain how to detect overflowed cells in Aspose.Cells before saving and adjust HtmlSaveOptions to control tooltip generation. | Show a post‑processing script that parses the generated HTML and removes tooltip attributes from cells whose text does not overflow. | Provide a step‑by‑step guide to customize tooltip behavior per cell in Aspose.Cells HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsTooltipDemo
{
    // Shows how to save a workbook to HTML with HtmlSaveOptions.AddTooltipText turned on, while ensuring that only cells whose text overflows the column width receive a tooltip attribute, leaving fitting cells tooltip‑free.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Cell A1 contains long text that exceeds the column width -> tooltip expected
            cells["A1"].PutValue("This is a very long text that will not fit into the column and should show a tooltip when exported to HTML.");

            // Cell B1 contains short text that fits within the column width -> no tooltip should be added
            cells["B1"].PutValue("Short");

            // Set a narrow column width so that A1 overflows but B1 fits
            cells.SetColumnWidth(0, 10); // Column A
            cells.SetColumnWidth(1, 10); // Column B

            // Configure HTML save options to enable tooltip generation
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.AddTooltipText = true; // Enable tooltip for overflowed cells

            // Save the workbook as HTML
            string outputPath = "TooltipDemo.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
