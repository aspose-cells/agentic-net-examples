// Title: Export Workbook to HTML with Print Area Only and Inline Styles (Aspose.Cells for .NET)
// Description: Shows how to create a workbook, set a print area (e.g., B2:F10), configure HtmlSaveOptions to export only that range and disable external CSS (using inline styles), then save the result as an HTML file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | HTML export | ExportPrintAreaOnly | DisableCss | inline styles | HtmlSaveOptions | print area | Excel to HTML
// Common Searches: Aspose.Cells export specific range to HTML | disable CSS when saving workbook as HTML Aspose.Cells | set print area before HTML export Aspose.Cells .NET | inline style HTML output Aspose.Cells | how to use HtmlSaveOptions ExportPrintAreaOnly
// Developer Intent: Save an Excel workbook as an HTML file that contains only the defined print area and uses inline styling instead of external CSS.
// Use Cases: Create a lightweight HTML preview of a selected worksheet range for web dashboards. | Generate email‑compatible HTML snippets without linking external style sheets. | Produce printable HTML sections from a larger workbook by defining a print area.
// AI Prompts: Modify the example to set a custom page orientation before exporting to HTML. | Show how to write the HTML output to a MemoryStream while keeping ExportPrintAreaOnly and DisableCss enabled. | Explain how to add a custom CSS class for table borders while preserving inline styles for other elements.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, set a print area (e.g., B2:F10), configure HtmlSaveOptions to export only that range and disable external CSS (using inline styles), then save the result as an HTML file with Aspose.Cells for .NET.
    public class ExportPrintAreaHtmlDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
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

                // Define the print area (e.g., B2:F10)
                worksheet.PageSetup.PrintArea = "B2:F10";

                // Set HTML save options:
                // - Export only the defined print area
                // - Disable CSS generation (use inline styles only)
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    ExportPrintAreaOnly = true,
                    DisableCss = true
                };

                // Save the workbook as an HTML file with the specified options
                string outputPath = "PrintArea_NoCss.html";
                workbook.Save(outputPath, options);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaHtmlDemo.Run();
        }
    }
}
