// Title: Export Print Area to HTML with Inline Styles – Aspose.Cells for .NET
// Description: Shows how to define a worksheet's print area (B2:F10) and save it as HTML using Aspose.Cells, with ExportPrintAreaOnly enabled and CSS generation disabled so the output contains only the selected range styled inline.
// Keywords: Aspose.Cells | C# HTML export | ExportPrintAreaOnly | DisableCss | print area HTML | inline styles | .NET | Workbook to HTML | save specific range | HtmlSaveOptions
// Common Searches: Aspose.Cells export specific range to HTML | How to disable CSS when saving workbook as HTML | Export print area only Aspose.Cells C# | HtmlSaveOptions ExportPrintAreaOnly example | Generate HTML preview of Excel range without external stylesheet
// Developer Intent: Create an HTML file that contains only a worksheet's print area with all styling applied inline, avoiding external CSS files.
// Use Cases: Embedding a concise Excel section in a web page without extra stylesheet files. | Sending a formatted table in an email where all styling must be inline. | Providing a lightweight printable preview of a report segment for browsers.
// AI Prompts: Show C# code to export a worksheet's print area to HTML using Aspose.Cells while turning off CSS files. | Give an example of HtmlSaveOptions with ExportPrintAreaOnly and DisableCss set to true. | Explain how to verify that the generated HTML includes only the defined range and uses inline styling.

using System;
using Aspose.Cells;

// Shows how to define a worksheet's print area (B2:F10) and save it as HTML using Aspose.Cells, with ExportPrintAreaOnly enabled and CSS generation disabled so the output contains only the selected range styled inline.
class ExportPrintAreaHtml
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

        // Define the print area that should be exported
        worksheet.PageSetup.PrintArea = "B2:F10";

        // Set up HTML save options:
        // - ExportPrintAreaOnly = true  => only the defined print area is saved
        // - DisableCss = true           => use only inline styles, no external CSS
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportPrintAreaOnly = true,
            DisableCss = true
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("PrintAreaOnly.html", htmlOptions);
    }
}
