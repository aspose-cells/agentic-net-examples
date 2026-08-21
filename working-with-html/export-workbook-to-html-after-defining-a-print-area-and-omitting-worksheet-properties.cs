// Title: Export Workbook to HTML with Print Area Only and No Worksheet Properties – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills a 20×10 grid, sets the print area to B2:F10, configures HtmlSaveOptions (ExportPrintAreaOnly = true, ExportWorksheetProperties = false) and saves the file as HTML, resulting in a lightweight page that contains only the selected range.
// Keywords: Aspose.Cells | C# | HTML export | print area | ExportPrintAreaOnly | ExportWorksheetProperties | HtmlSaveOptions | save workbook as HTML | worksheet metadata | range export
// Common Searches: Aspose.Cells export specific range to HTML | How to hide worksheet properties in HTML output Aspose.Cells | C# HtmlSaveOptions ExportPrintAreaOnly example | Save workbook as HTML without sheet metadata | Set print area before HTML export Aspose.Cells
// Developer Intent: Generate an HTML file that includes only the defined print area of a worksheet and omits all worksheet‑level properties.
// Use Cases: Display a compact HTML preview of a selected data block on a web portal. | Embed a small HTML snippet in email newsletters without extra worksheet metadata. | Create printable HTML sections from large workbooks while keeping file size minimal. | Provide a fast, metadata‑free HTML view for API consumers that need only a specific range.
// AI Prompts: Show how to set page orientation and margins before exporting the print area to HTML with Aspose.Cells. | Give an example that exports multiple worksheets, each with its own print area, to separate HTML files while disabling worksheet properties. | Explain how to programmatically retrieve the actual cell range that will be exported when ExportPrintAreaOnly is enabled.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills a 20×10 grid, sets the print area to B2:F10, configures HtmlSaveOptions (ExportPrintAreaOnly = true, ExportWorksheetProperties = false) and saves the file as HTML, resulting in a lightweight page that contains only the selected range.
    public class ExportPrintAreaWithoutWorksheetProperties
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (20 rows x 10 columns)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the print area (e.g., B2:F10)
            sheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Export only the defined print area
                ExportPrintAreaOnly = true,

                // Omit worksheet properties from the generated HTML
                ExportWorksheetProperties = false
            };

            // Save the workbook as HTML using the configured options
            string outputPath = "PrintAreaWithoutWorksheetProps.html";
            workbook.Save(outputPath, options);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
