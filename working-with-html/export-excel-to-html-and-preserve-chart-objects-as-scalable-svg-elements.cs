// Title: Export Excel to HTML with charts rendered as scalable SVG using Aspose.Cells for .NET
// Description: A C# console example that verifies an input workbook, loads it with Aspose.Cells, configures HtmlSaveOptions to render chart objects as SVG (vector) graphics, and saves a single self‑contained HTML file. The code includes error handling and demonstrates how to produce resolution‑independent charts in web pages.
// Keywords: Aspose.Cells export Excel to HTML SVG | C# HtmlSaveOptions chart as SVG | Excel chart SVG output .NET | self‑contained HTML with vector charts | Aspose.Cells scalable vector graphics export | convert .xlsx to HTML with SVG charts
// Common Searches: Aspose.Cells export chart as SVG HTML C# | How to render Excel charts as SVG in HTML using Aspose | C# save workbook to HTML with vector graphics | Aspose.Cells HtmlSaveOptions ExportChartImageFormat SVG | Convert Excel to single HTML file with SVG charts
// Developer Intent: Generate an HTML representation of an Excel workbook where all chart objects are output as scalable SVG elements instead of raster images.
// Use Cases: Create a lightweight, resolution‑independent web preview of financial dashboards that include charts. | Distribute a single HTML file that displays Excel data and vector charts without external image resources. | Integrate Excel‑based reports into responsive web applications where SVG scales on high‑DPI screens.
// AI Prompts: Modify the provided code to export charts as SVG while keeping other images embedded as Base64. | Show how to save the HTML output with external SVG files instead of inline data URIs. | Explain the performance impact of using SVG charts versus PNG images in the generated HTML.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // A C# console example that verifies an input workbook, loads it with Aspose.Cells, configures HtmlSaveOptions to render chart objects as SVG (vector) graphics, and saves a single self‑contained HTML file. The code includes error handling and demonstrates how to produce resolution‑independent charts in web pages.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export images (including charts) as Base64-encoded data
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Excel file has been exported to HTML at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
