// Title: Export Worksheet Print Area to HTML without Document Properties using Aspose.Cells for .NET
// Description: Shows how to define a print area (B2:F10), set HtmlSaveOptions.ExportPrintAreaOnly = true and ExportDocumentProperties = false, and save the workbook as a compact HTML file.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportPrintAreaOnly | ExportDocumentProperties | print area export | HTML spreadsheet export | remove workbook metadata | lightweight HTML preview
// Common Searches: Aspose.Cells export specific range to HTML | C# save worksheet as HTML without document properties | Export print area only using HtmlSaveOptions | How to hide workbook metadata in HTML export Aspose.Cells | Generate HTML preview of printable area in .NET
// Developer Intent: Create an HTML file that contains only the worksheet's defined print area and excludes all document properties.
// Use Cases: Provide a fast HTML preview that displays only the printable section of a large spreadsheet. | Produce HTML reports that comply with data‑privacy policies by removing workbook metadata. | Reduce file size for web‑based spreadsheet viewers by exporting a specific range only.
// AI Prompts: Give me C# code to export a worksheet's print area to HTML with Aspose.Cells while omitting document properties. | How do I configure HtmlSaveOptions to output only the defined print range and exclude metadata? | Explain the impact of ExportPrintAreaOnly and ExportDocumentProperties on the generated HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaExport
{
    // Shows how to define a print area (B2:F10), set HtmlSaveOptions.ExportPrintAreaOnly = true and ExportDocumentProperties = false, and save the workbook as a compact HTML file.
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
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the print area that should be exported (e.g., B2:F10)
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options:
            // - ExportPrintAreaOnly = true  => only the defined print area is exported.
            // - ExportDocumentProperties = false => document properties are omitted from the output.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true,
                ExportDocumentProperties = false
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("PrintAreaOnly_NoDocProps.html", options);
        }
    }
}
