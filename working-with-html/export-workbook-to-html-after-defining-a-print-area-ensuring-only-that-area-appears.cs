// Title: Export a Defined Print Area to HTML with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills cells, sets a print area (e.g., B2:F10), configures HtmlSaveOptions to export only that region, and saves the result as an HTML file containing the selected range.
// Keywords: Aspose.Cells | C# | .NET | HTML export | print area | ExportPrintAreaOnly | worksheet range to HTML | selective HTML save | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells export only print area to HTML C# | HtmlSaveOptions ExportPrintAreaOnly example | Save specific cell range as HTML using Aspose.Cells | C# generate HTML from worksheet print area | Aspose.Cells limit HTML output to a range
// Developer Intent: Produce an HTML file that includes just the worksheet's defined print area using Aspose.Cells for .NET.
// Use Cases: Provide a clean HTML preview of a targeted data block for web reporting. | Embed a selected spreadsheet segment in a dashboard or email without extra rows/columns. | Create lightweight printable HTML snippets for a specific range of cells.
// AI Prompts: Write C# code with Aspose.Cells to export only the print area B2:F10 to HTML. | Show how to include grid lines when exporting a defined print area to HTML using Aspose.Cells. | Explain how to calculate a dynamic print area based on content and save it as HTML.

using System;
using Aspose.Cells;

// Creates a workbook, fills cells, sets a print area (e.g., B2:F10), configures HtmlSaveOptions to export only that region, and saves the result as an HTML file containing the selected range.
class ExportPrintAreaToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
            }
        }

        // Define the print area that should be exported (e.g., B2:F10)
        worksheet.PageSetup.PrintArea = "B2:F10";

        // Configure HTML save options to export only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportPrintAreaOnly = true;   // Export only the print area
        // htmlOptions.ExportGridLines = true;    // Optional: include grid lines in the HTML

        // Save the workbook as an HTML file; only the print area will appear in the output
        workbook.Save("PrintAreaOnly.html", htmlOptions);
    }
}
