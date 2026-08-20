// Title: Export Worksheet Print Area to HTML with Aspose.Cells (C#)
// Description: Demonstrates how to define a print area (e.g., B2:F10) in an Aspose.Cells workbook, configure HtmlSaveOptions with ExportPrintAreaOnly = true (and optional grid lines), and save the selected range as a lightweight HTML file.
// Keywords: Aspose.Cells C# HTML export | ExportPrintAreaOnly | print area to HTML | HtmlSaveOptions grid lines | save selected range as HTML
// Common Searches: Aspose.Cells export only print area to HTML | C# HtmlSaveOptions ExportPrintAreaOnly example | How to save a specific range as HTML with Aspose.Cells | Include grid lines when exporting HTML from Aspose.Cells
// Developer Intent: Generate an HTML file that contains only the worksheet's defined print area.
// Use Cases: Create a web‑ready snapshot of a report section for dashboards. | Produce compact HTML snippets for email attachments by limiting output to the print area. | Render printable HTML views of invoices or forms while preserving cell borders.
// AI Prompts: Show me a C# snippet that sets a print area and exports only that range to HTML using Aspose.Cells. | How can I include grid lines when saving a workbook as HTML with ExportPrintAreaOnly enabled? | Explain how to export multiple print areas or adjust page setup before converting to HTML with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to define a print area (e.g., B2:F10) in an Aspose.Cells workbook, configure HtmlSaveOptions with ExportPrintAreaOnly = true (and optional grid lines), and save the selected range as a lightweight HTML file.
    public class ExportPrintAreaToHtml
    {
        public static void Run()
        {
            try
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

                // Define the print area (e.g., B2:F10)
                worksheet.PageSetup.PrintArea = "B2:F10";

                // Configure HTML save options to export only the defined print area
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportPrintAreaOnly = true, // Export only the print area
                    ExportGridLines = true      // Optional: include grid lines in the output
                };

                // Save the workbook as HTML; only the print area will be exported
                workbook.Save("PrintAreaOnly.html", htmlOptions);
                Console.WriteLine("HTML file 'PrintAreaOnly.html' has been created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaToHtml.Run();
        }
    }
}
