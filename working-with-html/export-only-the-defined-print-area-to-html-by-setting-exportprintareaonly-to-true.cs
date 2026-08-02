// Title: Export Worksheet Print Area to HTML with Aspose.Cells for .NET (ExportPrintAreaOnly)
// Description: Creates a workbook, fills cells with sample data, defines a print area (e.g., B2:F10), sets HtmlSaveOptions.ExportPrintAreaOnly to true (optionally includes grid lines), and saves the result as PrintAreaOnly.html. The generated HTML contains only the specified print range.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportPrintAreaOnly | C# | .NET | print area HTML export | worksheet to HTML | grid lines | sample code | save as HTML
// Common Searches: Aspose.Cells export only print area to HTML | HtmlSaveOptions ExportPrintAreaOnly C# example | How to save a specific range as HTML with Aspose.Cells | Include grid lines when exporting Excel to HTML Aspose | Export selected cells to HTML using Aspose.Cells .NET
// Developer Intent: Generate an HTML file that includes only the worksheet’s defined print area.
// Use Cases: Provide a web‑ready view of a report section without extra rows/columns. | Embed a focused data block in a dashboard or email by exporting just the print range. | Create a printable HTML snippet for a specific area of a spreadsheet.
// AI Prompts: Write C# code that automatically sets the print area to the used range and exports it to HTML with grid lines using Aspose.Cells. | Explain how ExportPrintAreaOnly influences the HTML output and list other HtmlSaveOptions that can be combined for customized exports. | Show how to loop through multiple worksheets, set individual print areas, and save each as a separate HTML file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace ExportPrintAreaToHtml
{
    // Creates a workbook, fills cells with sample data, defines a print area (e.g., B2:F10), sets HtmlSaveOptions.ExportPrintAreaOnly to true (optionally includes grid lines), and saves the result as PrintAreaOnly.html. The generated HTML contains only the specified print range.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
                }
            }

            // Define the print area (e.g., B2:F10)
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options to export only the defined print area
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportPrintAreaOnly = true;   // Export only the print area
            htmlOptions.ExportGridLines = true;       // Optional: include grid lines in the output

            // Save the workbook as HTML
            workbook.Save("PrintAreaOnly.html", htmlOptions);
        }
    }
}
