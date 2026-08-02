// Title: Export Excel Gridlines to HTML with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, sets worksheet.IsGridlinesVisible to true, configures HtmlSaveOptions with ExportGridLines (and optionally ExportActiveWorksheetOnly), and saves the result as an HTML file that preserves the original gridlines.
// Keywords: Aspose.Cells | C# | ExportGridLines | HtmlSaveOptions | Excel to HTML | gridlines | worksheet.IsGridlinesVisible | export active worksheet | HTML preview | Excel web rendering
// Common Searches: Aspose.Cells export gridlines to HTML | C# save Excel as HTML with gridlines | keep Excel gridlines in HTML output | export only active worksheet to HTML Aspose | HtmlSaveOptions ExportGridLines example
// Developer Intent: Generate an HTML file from an Excel workbook that displays the worksheet’s gridlines.
// Use Cases: Provide a web‑based preview of an Excel sheet that retains cell borders for clearer readability. | Create downloadable HTML reports that maintain the original spreadsheet layout. | Embed a single worksheet in a web page while preserving its gridline formatting.
// AI Prompts: Write C# code using Aspose.Cells to export an Excel worksheet to HTML with gridlines and only the active sheet. | Explain the relationship between worksheet.IsGridlinesVisible and HtmlSaveOptions.ExportGridLines when saving to HTML. | Show how to export multiple worksheets to separate HTML files, each with gridlines enabled, using Aspose.Cells.

using System;
using Aspose.Cells;

namespace ExportGridLinesToHtml
{
    // Loads an Excel workbook, sets worksheet.IsGridlinesVisible to true, configures HtmlSaveOptions with ExportGridLines (and optionally ExportActiveWorksheetOnly), and saves the result as an HTML file that preserves the original gridlines.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can modify this as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure gridlines are visible in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Create HTML save options and enable gridline export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Export the gridlines to HTML
                ExportActiveWorksheetOnly = true      // Export only the active worksheet (optional)
            };

            // Save the workbook as an HTML file with gridlines displayed
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with gridlines exported.");
        }
    }
}
