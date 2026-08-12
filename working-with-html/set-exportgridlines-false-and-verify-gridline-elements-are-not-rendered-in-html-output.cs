// Title: Export HTML without Gridlines using Aspose.Cells for .NET and Verify the Output
// Description: Demonstrates how to disable gridline export with HtmlSaveOptions (ExportGridLines = false), save only the active worksheet as HTML, and programmatically confirm that no gridline or border markup appears in the generated file.
// Keywords: Aspose.Cells HTML export | ExportGridLines false | disable gridlines .NET | HtmlSaveOptions gridlines | verify HTML output Aspose | C# spreadsheet to HTML | Aspose.Cells testing | web report without gridlines | Aspose.Cells .NET Europe | Aspose.Cells US developers
// Common Searches: Aspose.Cells ExportGridLines example C# | How to hide gridlines when saving to HTML with Aspose.Cells | Check HTML output for gridline CSS using Aspose.Cells | Aspose.Cells HtmlSaveOptions without borders | Validate that gridlines are not exported in HTML
// Developer Intent: Save a workbook as HTML without rendering gridlines and programmatically ensure the resulting markup contains no gridline or border elements.
// Use Cases: Create clean, border‑free HTML reports from Excel data for web publishing. | Apply custom CSS styling to worksheets when default gridlines are unnecessary. | Automated unit tests that verify the ExportGridLines setting suppresses gridline markup.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to HTML with ExportGridLines set to false and then checks the file for any "gridline" or "border" strings. | Explain step‑by‑step how to programmatically confirm that gridline CSS is absent from the HTML saved by Aspose.Cells. | Suggest alternative verification techniques (e.g., DOM parsing, regex) to ensure gridlines are omitted in the exported HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsGridlinesDemo
{
    // Demonstrates how to disable gridline export with HtmlSaveOptions (ExportGridLines = false), save only the active worksheet as HTML, and programmatically confirm that no gridline or border markup appears in the generated file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data so the worksheet has visible content
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue("Row 1");
            worksheet.Cells["A3"].PutValue("Row 2");

            // Ensure gridlines are visible in the worksheet (default is true)
            worksheet.IsGridlinesVisible = true;

            // Configure HTML save options to NOT export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = false,               // Disable gridline export
                ExportActiveWorksheetOnly = true       // Export only the first worksheet
            };

            // Save the workbook as HTML
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            // Load the generated HTML file for verification
            string htmlContent = File.ReadAllText(htmlPath);

            // Simple verification: check for typical gridline CSS class/attributes
            bool containsGridlineMarkers = htmlContent.Contains("gridline") ||
                                           htmlContent.Contains("border");

            // Output verification result
            if (containsGridlineMarkers)
                Console.WriteLine("Gridlines were found in the HTML output (unexpected).");
            else
                Console.WriteLine("Gridlines are not present in the HTML output as expected.");
        }
    }
}
