// Title: Aspose.Cells C# – Export Workbook to HTML without Gridlines and Verify
// Description: Creates a workbook, adds sample data, enables worksheet gridlines, configures HtmlSaveOptions with ExportGridLines = false, saves to HTML, reads the file, and checks for border CSS to confirm that gridlines are omitted.
// Keywords: Aspose.Cells HTML export C# | ExportGridLines false | disable gridlines Aspose.Cells | verify HTML output Aspose | gridline CSS detection | Aspose.Cells HtmlSaveOptions example
// Common Searches: How to turn off gridlines when saving Excel to HTML with Aspose.Cells | Aspose.Cells ExportGridLines property C# example | Check HTML for gridline borders after Aspose export | Validate that Aspose.Cells HTML export excludes gridlines
// Developer Intent: Save a workbook as HTML with gridlines turned off and programmatically confirm that no gridline markup appears.
// Use Cases: Generate clean HTML reports where spreadsheet borders are not wanted. | Create printable web versions of worksheets without visual clutter. | Automate CI tests to ensure ExportGridLines setting is respected.
// AI Prompts: Write C# code that exports an Aspose.Cells workbook to HTML with ExportGridLines set to false and validates the absence of gridline CSS. | Explain how to scan the generated HTML file to detect any border styles that would indicate exported gridlines. | Suggest alternative verification techniques, such as DOM parsing or regex, to confirm gridlines are omitted.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates a workbook, adds sample data, enables worksheet gridlines, configures HtmlSaveOptions with ExportGridLines = false, saves to HTML, reads the file, and checks for border CSS to confirm that gridlines are omitted.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data so the HTML has visible cells
            sheet.Cells["A1"].PutValue("First");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Ensure gridlines are visible in the worksheet (they would be exported if enabled)
            sheet.IsGridlinesVisible = true;

            // Configure HTML save options with ExportGridLines disabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = false,               // Disable gridline export
                ExportActiveWorksheetOnly = true       // Export only the active sheet for simplicity
            };

            // Define output path
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            // Verify that gridlines are absent by inspecting the generated HTML
            string htmlContent = File.ReadAllText(outputPath);

            // Simple check: look for CSS border definitions that Aspose adds for gridlines
            bool containsGridLines = htmlContent.Contains("border") && htmlContent.Contains("solid");

            Console.WriteLine("HTML file saved to: " + Path.GetFullPath(outputPath));
            Console.WriteLine("ExportGridLines option set to: " + htmlOptions.ExportGridLines);
            Console.WriteLine("Gridlines present in HTML? " + (containsGridLines ? "Yes" : "No"));
            Console.WriteLine(containsGridLines
                ? "Gridlines were exported despite the option being disabled."
                : "Gridlines are correctly absent in the HTML output.");
        }
    }
}
