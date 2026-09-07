// Title: Export the first chart from an Excel workbook to an SVG file with selectable text using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, retrieves the first chart on the first worksheet, and saves it as an SVG while keeping the text selectable with Aspose.Cells. | Show how to configure ImageOrPrintOptions.OnePagePerSheet for SVG chart export so the resulting SVG preserves searchable text. | Write a C# routine that loops through all charts in a worksheet and exports each chart to its own SVG file using Aspose.Cells.
// Common Searches: aspnet export excel chart to svg with selectable text using aspose.cells | c# how to save a chart as scalable vector graphics from a workbook | imageorprintoptions onepagepersheet svg chart export example | export multiple charts from a worksheet to separate svg files asp.net | preserve text layer when converting excel chart to svg asp.net
// Tags: chart export to SVG Aspose.Cells | ImageOrPrintOptions OnePagePerSheet SVG | selectable text in SVG chart Aspose | C# export Excel chart as scalable vector graphics | multiple chart SVG export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, accesses the first worksheet's first chart, and exports it to an SVG file using Aspose.Cells with ImageOrPrintOptions.OnePagePerSheet to retain selectable text.
class ExportChartToSvg
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.svg";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Configure image options for SVG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat property is optional; SVG is inferred from file extension
                OnePagePerSheet = true // Required flag for image export
            };

            // Export the chart to an SVG file
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
