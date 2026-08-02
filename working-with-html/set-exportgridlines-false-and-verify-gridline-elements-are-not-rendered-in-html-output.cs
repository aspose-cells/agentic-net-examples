// Title: Aspose.Cells for .NET – Disable Gridlines When Exporting a Workbook to HTML
// Description: Demonstrates how to set HtmlSaveOptions.ExportGridLines to false, save only the active worksheet as HTML, and programmatically confirm that no grid‑line markup appears in the output.
// Keywords: Aspose.Cells ExportGridLines false | disable gridlines HTML export | HtmlSaveOptions C# | verify HTML output Aspose.Cells | C# export workbook without gridlines
// Common Searches: Aspose.Cells prevent gridlines in HTML | HtmlSaveOptions ExportGridLines example C# | how to hide gridlines when saving to HTML Aspose.Cells | check for gridline CSS in exported HTML | export only active worksheet Aspose.Cells
// Developer Intent: Turn off grid‑line rendering in the HTML file generated from an Aspose.Cells workbook.
// Use Cases: Create clean web‑ready HTML reports without visible spreadsheet gridlines. | Automated validation that exported HTML does not contain grid‑line CSS classes. | Generate single‑sheet HTML snippets for emails or documentation while suppressing gridlines.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with ExportGridLines set to false and verify that the output lacks any "gridline" CSS. | Create a C# unit test that enables worksheet gridlines, saves to HTML with ExportGridLines disabled, and asserts the HTML contains no gridline markup. | Explain how the ExportGridLines property affects HTML output in Aspose.Cells and describe methods to validate its effect.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportGridLinesDemo
{
    // Demonstrates how to set HtmlSaveOptions.ExportGridLines to false, save only the active worksheet as HTML, and programmatically confirm that no grid‑line markup appears in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable gridlines in the worksheet (so they would be exported if not disabled)
            sheet.IsGridlinesVisible = true;

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["A3"].PutValue("Row 2");

            // Configure HTML save options: explicitly disable exporting gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = false,               // Do not export gridlines
                ExportActiveWorksheetOnly = true       // Export only the first worksheet
            };

            // Define output HTML file path
            string htmlPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            // Verify that gridline related markup is not present in the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);

            // Simple verification: check for the presence of the CSS class used for gridlines
            // Aspose.Cells typically uses "gridline" in the style definitions.
            bool containsGridlines = htmlContent.IndexOf("gridline", StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine("ExportGridLines set to: " + htmlOptions.ExportGridLines);
            Console.WriteLine("Gridline markup found in HTML: " + containsGridlines);
            Console.WriteLine(containsGridlines
                ? "Gridlines were exported unexpectedly."
                : "Gridlines were not exported as expected.");
        }
    }
}
