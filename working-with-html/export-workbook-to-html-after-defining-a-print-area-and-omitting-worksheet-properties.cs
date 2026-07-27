// Title: Export Print Area to HTML without Worksheet Properties using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills a 10×6 grid, sets the print area to B2:F10, configures HtmlSaveOptions to exclude worksheet properties and to export only the defined print area, then saves the result as an HTML file.
// Keywords: Aspose.Cells | C# | HTML export | print area | ExportWorksheetProperties | ExportPrintAreaOnly | save workbook as HTML | omit worksheet metadata | HtmlSaveOptions | export specific range to HTML
// Common Searches: Aspose.Cells export specific range to HTML | How to hide worksheet properties in HTML output Aspose | C# HtmlSaveOptions ExportPrintAreaOnly example | Save Excel as HTML without sheet metadata | Aspose.Cells print area only HTML
// Developer Intent: Generate an HTML file that contains only the cells inside a defined print area and excludes all worksheet‑level properties and metadata.
// Use Cases: Web preview of a report section by exporting only the relevant range. | Embedding a data slice in email attachments without extra worksheet information. | Creating lightweight HTML for dashboards or portals while keeping file size minimal. | Producing printable HTML for a specific area of a spreadsheet.
// AI Prompts: Write C# code using Aspose.Cells to export only the print area B2:F10 to HTML and suppress worksheet properties. | Explain how ExportWorksheetProperties and ExportPrintAreaOnly affect the generated HTML file. | Show how to add custom CSS styling to the exported HTML range. | Demonstrate exporting multiple worksheets, each with its own print area, to separate HTML files.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills a 10×6 grid, sets the print area to B2:F10, configures HtmlSaveOptions to exclude worksheet properties and to export only the defined print area, then saves the result as an HTML file.
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (10 rows x 6 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 6; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the print area (B2:F10)
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options:
            // - Omit worksheet properties from the output
            // - Export only the defined print area
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportWorksheetProperties = false,
                ExportPrintAreaOnly = true // comment out if you want the whole sheet
            };

            // Save the workbook as HTML
            string outputPath = "PrintAreaWithoutWorksheetProps.html";
            workbook.Save(outputPath, options);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
