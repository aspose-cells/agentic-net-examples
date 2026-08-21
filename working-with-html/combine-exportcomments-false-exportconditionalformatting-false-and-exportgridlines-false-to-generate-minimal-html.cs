// Title: Generate Minimal HTML from an Excel Workbook with Aspose.Cells .NET – Disable Comments, Conditional Formatting & Gridlines
// Description: C# example that creates or loads a workbook, sets HtmlSaveOptions.IsExportComments, ExportConditionalFormatting and ExportGridLines to false, and saves a lightweight HTML file containing only cell values.
// Keywords: Aspose.Cells minimal HTML export | disable comments Aspose.Cells | remove conditional formatting HTML | export without gridlines | lightweight Excel to HTML .NET
// Common Searches: Aspose.Cells export minimal HTML | how to hide comments when saving Excel as HTML | disable conditional formatting in HTML output Aspose | remove grid lines from HTML export Aspose.Cells | C# generate lightweight HTML from workbook
// Developer Intent: Produce an HTML representation of a workbook that includes only raw cell data, omitting comments, conditional formatting rules, and grid lines.
// Use Cases: Create clean HTML reports for web dashboards without extra styling artifacts. | Generate compact HTML email bodies from spreadsheets, keeping file size low. | Export data‑only views of Excel sheets for documentation or API responses.
// AI Prompts: Show C# code using Aspose.Cells to save a workbook as minimal HTML with comments, conditional formatting, and grid lines disabled. | Explain the impact of IsExportComments, ExportConditionalFormatting, and ExportGridLines on the size and appearance of the generated HTML. | Provide a step‑by‑step guide to load an existing .xlsx file and export it to minimal HTML using Aspose.Cells .NET.

using System;
using Aspose.Cells;

namespace MinimalHtmlExport
{
    // C# example that creates or loads a workbook, sets HtmlSaveOptions.IsExportComments, ExportConditionalFormatting and ExportGridLines to false, and saves a lightweight HTML file containing only cell values.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // creates a new workbook

                // Add some sample data (optional, just to have content)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B2"].PutValue(123);

                // Configure HTML save options for minimal output
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do not export comments
                    IsExportComments = false,

                    // Do not export grid lines
                    ExportGridLines = false
                };

                // Save the workbook as HTML with the specified options
                workbook.Save("minimal_output.html", htmlOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
