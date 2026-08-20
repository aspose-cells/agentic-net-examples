// Title: Export a Print Area to HTML with Inline Styles Using Aspose.Cells for .NET
// Description: This example creates a workbook, fills it with sample data, defines a print area (B2:F10), and saves the sheet as an HTML file. The HtmlSaveOptions are configured to export only the specified range and to embed all styling inline, eliminating external CSS files.
// Keywords: Aspose.Cells | C# HTML export | ExportPrintAreaOnly | DisableCss | print area HTML | inline styles | HtmlSaveOptions | save workbook as HTML | no external CSS | range export
// Common Searches: Aspose.Cells export specific range to HTML | How to disable CSS when saving workbook as HTML in .NET | Export print area only Aspose.Cells | C# generate HTML from Excel without external stylesheet | Save Excel range as HTML with inline styles
// Developer Intent: Generate an HTML file that contains only the defined print area and uses inline styling instead of external CSS.
// Use Cases: Create a lightweight HTML preview of a selected spreadsheet region for email newsletters. | Embed a spreadsheet snippet in a web page without loading additional CSS resources. | Produce a printable HTML section for documentation or reporting tools.
// AI Prompts: Show how to add a custom page header to the HTML output while keeping ExportPrintAreaOnly and DisableCss enabled. | Provide code that writes the HTML to a MemoryStream instead of a file, preserving the print‑area‑only and inline‑style settings. | Explain how to embed images as base64 data URIs in the HTML while still disabling external CSS.

using System;
using Aspose.Cells;

// This example creates a workbook, fills it with sample data, defines a print area (B2:F10), and saves the sheet as an HTML file. The HtmlSaveOptions are configured to export only the specified range and to embed all styling inline, eliminating external CSS files.
class ExportPrintAreaHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data
        for (int i = 0; i < 20; i++)
        {
            for (int j = 0; j < 10; j++)
            {
                worksheet.Cells[i, j].PutValue($"Cell {i + 1},{j + 1}");
            }
        }

        // Define the print area that will be exported
        worksheet.PageSetup.PrintArea = "B2:F10";

        // Set HTML save options: export only the defined print area and disable CSS generation
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.ExportPrintAreaOnly = true; // export only the print area
        options.DisableCss = true;          // use only inline styles, no external CSS

        // Save the workbook as an HTML file
        workbook.Save("PrintArea_NoCss.html", options);
    }
}
