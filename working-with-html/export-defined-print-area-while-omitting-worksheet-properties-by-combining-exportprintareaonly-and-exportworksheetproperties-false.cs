// Title: Export a Defined Print Area to HTML without Worksheet Properties – Aspose.Cells for .NET
// Description: Creates a workbook, sets the print area to B2:F10, and saves it as HTML using HtmlSaveOptions with ExportPrintAreaOnly enabled and ExportWorksheetProperties disabled, producing a lightweight HTML file that contains only the selected range and no worksheet metadata.
// Keywords: Aspose.Cells HTML export | ExportPrintAreaOnly | ExportWorksheetProperties false | C# print area to HTML | Aspose.Cells save as HTML | grid lines HTML export | remove worksheet metadata HTML
// Common Searches: Aspose.Cells export only print area to HTML | HTML export without worksheet properties Aspose | C# save specific range as HTML Aspose.Cells | how to hide workbook metadata in HTML export | export B2:F10 to HTML using Aspose
// Developer Intent: Generate an HTML file that includes only the defined print area of a worksheet and excludes all worksheet property information.
// Use Cases: Create a compact HTML preview of a report section for web publishing. | Embed a specific table range in an email template without exposing workbook settings. | Render a dashboard widget with grid lines while keeping the HTML payload minimal.
// AI Prompts: Show C# code to export only the print area to HTML with Aspose.Cells, disabling worksheet properties. | How do I set ExportPrintAreaOnly and ExportWorksheetProperties in HtmlSaveOptions for a .NET workbook? | Explain the steps to generate HTML that contains only cells B2:F10 and omits page setup data.

using System;
using Aspose.Cells;

namespace ExportPrintAreaOnlyDemo
{
    // Creates a workbook, sets the print area to B2:F10, and saves it as HTML using HtmlSaveOptions with ExportPrintAreaOnly enabled and ExportWorksheetProperties disabled, producing a lightweight HTML file that contains only the selected range and no worksheet metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
                }
            }

            // Define the print area that should be exported
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Set HTML save options:
            // - ExportPrintAreaOnly = true  => only the defined print area is exported
            // - ExportWorksheetProperties = false => worksheet properties are omitted from the HTML
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true,
                ExportWorksheetProperties = false,
                ExportGridLines = true // optional, keeps grid lines visible
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("PrintAreaOnly.html", options);
        }
    }
}
