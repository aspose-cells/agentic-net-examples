// Title: Export a worksheet's print area to HTML with inline styles using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set a print area (B2:F10) in an Aspose.Cells workbook, configure HtmlSaveOptions to export only that range (ExportPrintAreaOnly = true) and suppress external CSS (DisableCss = true), and save the result as a lightweight HTML file with inline styling.
// Keywords: Aspose.Cells HTML export | ExportPrintAreaOnly | DisableCss | inline styles | C# Excel to HTML | print area only | Aspose.Cells .NET example | HTMLSaveOptions | Excel print area HTML | no external CSS Aspose
// Common Searches: Aspose.Cells export only print area to HTML | HtmlSaveOptions DisableCss C# example | ExportPrintAreaOnly true Aspose.Cells | how to generate HTML without CSS using Aspose.Cells | C# convert Excel range to HTML inline styles
// Developer Intent: Generate an HTML file that contains only the defined print area and uses inline styling, avoiding external CSS files.
// Use Cases: Creating compact HTML reports that show a specific cell range. | Embedding Excel data in email bodies where external CSS is blocked. | Producing printable web previews that match the Excel print layout. | Building lightweight web dashboards that load only necessary cells.
// AI Prompts: Provide C# code with Aspose.Cells to export a worksheet's print area to HTML using only inline styles. | Show how to set HtmlSaveOptions.DisableCss = true and ExportPrintAreaOnly = true in Aspose.Cells. | Explain the steps to define a print area and save it as CSS‑free HTML with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaExport
{
    // Demonstrates how to set a print area (B2:F10) in an Aspose.Cells workbook, configure HtmlSaveOptions to export only that range (ExportPrintAreaOnly = true) and suppress external CSS (DisableCss = true), and save the result as a lightweight HTML file with inline styling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (optional, demonstrates the print area)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the print area that will be exported
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options:
            // - ExportPrintAreaOnly = true  => only the defined print area is saved
            // - DisableCss = true           => use only inline styles, no external CSS
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportPrintAreaOnly = true;
            options.DisableCss = true;

            // Save the workbook as HTML using the configured options
            workbook.Save("PrintAreaOnly.html", options);
        }
    }
}
