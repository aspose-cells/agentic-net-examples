// Title: Aspose.Cells .NET – Save Workbook as HTML with AddTooltipText to Show Full Cell Text on Hover
// Description: Demonstrates how to create a workbook, insert a long string, narrow a column to force truncation, and export the sheet to HTML using HtmlSaveOptions with AddTooltipText enabled, so the complete cell value appears as a tooltip when the user hovers over the truncated cell.
// Keywords: Aspose.Cells | HtmlSaveOptions | AddTooltipText | tooltip | hover tooltip | truncated cell | C# | .NET | export to HTML | column width | cell text tooltip | workbook save HTML
// Common Searches: Aspose.Cells enable tooltip for HTML export | AddTooltipText property example C# | show full cell value on hover Aspose.Cells HTML | HTML export truncated cell tooltip Aspose | how to save workbook as HTML with tooltip Aspose.Cells
// Developer Intent: Generate HTML from a workbook where cells that exceed column width display their full content as a hover tooltip.
// Use Cases: Financial dashboards with narrow description columns that reveal full notes on mouseover. | Web‑based spreadsheet previews where space is limited but users need access to complete data. | Reporting portals that export Excel data to HTML while preserving readability through tooltips.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with AddTooltipText set to true and a narrow column to trigger truncation. | Explain the effect of HtmlSaveOptions.AddTooltipText on the generated HTML and which HTML attributes are added for tooltips. | Provide a step‑by‑step guide to ensure the output directory exists before saving the HTML file with tooltip support.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert a long string, narrow a column to force truncation, and export the sheet to HTML using HtmlSaveOptions with AddTooltipText enabled, so the complete cell value appears as a tooltip when the user hovers over the truncated cell.
    public class HtmlSaveWithTooltipDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell that will exceed the column width
            worksheet.Cells["A1"].PutValue("This is a very long text that will not fit into the cell width and should appear as a tooltip when hovered.");

            // Set a narrow column width to force truncation in the HTML view
            worksheet.Cells.SetColumnWidth(0, 10);

            // Configure HTML save options to enable tooltip text for truncated data
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                AddTooltipText = true // Enable tooltip
            };

            // Define output path and ensure the directory exists
            string outputPath = "output_with_tooltip.html";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("HTML file saved with tooltip enabled: " + outputPath);
        }
    }
}
