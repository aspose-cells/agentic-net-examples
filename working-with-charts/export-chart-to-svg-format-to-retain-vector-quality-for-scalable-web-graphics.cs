// Title: Export the first chart from an Excel workbook to a scalable SVG file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, retrieves the first worksheet chart, and saves it as an SVG image using ImageOrPrintOptions. | Show how to add robust error handling for missing workbook files or absent charts before exporting a chart to SVG. | Demonstrate configuring ImageOrPrintOptions for vector output when calling Chart.ToImage to produce an SVG file.
// Common Searches: Aspose.Cells C# export chart to SVG vector format | How to save Excel chart as SVG using Aspose.Cells .NET library | C# example converting worksheet chart to scalable SVG with Aspose.Cells | Export first chart from workbook to SVG file Aspose.Cells | ImageOrPrintOptions SVG output for chart Aspose.Cells .NET
// Tags: chart to SVG export Aspose.Cells | ImageOrPrintOptions SVG rendering C# | Excel chart vector graphic conversion | Aspose.Cells chart image generation | save worksheet chart as scalable SVG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads an .xlsx workbook, verifies that the first worksheet contains at least one chart, and uses Aspose.Cells' ImageOrPrintOptions with Chart.ToImage to export the first chart as a vector SVG file, including error handling for missing files or absent charts.
class ExportChartToSvg
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "chart.svg";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains a chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Set image options (no need to set ImageFormat explicitly; extension determines format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Export the chart to an SVG file
            chart.ToImage(outputPath, imgOptions);
            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
