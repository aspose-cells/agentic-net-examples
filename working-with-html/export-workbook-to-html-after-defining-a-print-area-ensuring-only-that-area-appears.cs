// Title: Export Defined Print Area to HTML with Aspose.Cells (C#)
// Description: Creates a workbook, sets a print area (e.g., B2:F10), configures HtmlSaveOptions with ExportPrintAreaOnly = true (and optional grid lines), and saves the file as HTML so that only the specified range appears in the output.
// Keywords: Aspose.Cells | C# | .NET | HTML export | print area | ExportPrintAreaOnly | HtmlSaveOptions | grid lines | worksheet preview
// Common Searches: Aspose.Cells export only print area to HTML | HtmlSaveOptions ExportPrintAreaOnly C# example | Save worksheet as HTML with grid lines Aspose.Cells | How to set print area before HTML export Aspose.Cells .NET
// Developer Intent: Generate an HTML file that contains only the worksheet’s defined print area.
// Use Cases: Show a web preview of a report section without loading the whole sheet. | Create a printable HTML snippet for a dashboard range. | Provide customers a lightweight HTML view of selected data while preserving grid lines.
// AI Prompts: Give me C# code that sets a print area in an Aspose.Cells worksheet and exports it to HTML using HtmlSaveOptions with ExportPrintAreaOnly enabled. | Explain how to export only the defined print area to HTML and include grid lines with Aspose.Cells for .NET. | Show a step‑by‑step guide to customize the HTML output (styles, images) when exporting a specific print area using Aspose.Cells.

using System;
using Aspose.Cells;

namespace ExportPrintAreaToHtml
{
    // Creates a workbook, sets a print area (e.g., B2:F10), configures HtmlSaveOptions with ExportPrintAreaOnly = true (and optional grid lines), and saves the file as HTML so that only the specified range appears in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (optional, just for demonstration)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
                }
            }

            // Define the print area that should be exported
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options to export only the defined print area
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true,   // Export only the print area
                ExportGridLines = true        // Optional: include grid lines in the HTML
            };

            // Save the workbook as HTML; only the print area will appear in the output file
            workbook.Save("PrintAreaOnly.html", htmlOptions);
        }
    }
}
