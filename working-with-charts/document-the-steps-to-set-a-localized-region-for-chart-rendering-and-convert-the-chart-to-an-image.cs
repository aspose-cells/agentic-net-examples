// Title: How to set a specific locale for chart rendering and export an Excel chart to PNG using Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that sets the workbook's CultureInfo to a target locale, renders a chart with that locale applied, and saves the chart as a PNG using ImageOrPrintOptions. | Show how to adjust regional settings such as number formats and axis labels for an Aspose.Cells chart before converting it to an image. | Generate a snippet that exports an Aspose.Cells chart to JPEG with a custom DPI while preserving the previously configured locale.
// Common Searches: Aspose.Cells C# set culture info for chart rendering before image export | Export Excel chart to PNG with specific locale using Aspose.Cells .NET | How to change regional settings of a chart in Aspose.Cells for .NET | C# Aspose.Cells chart to image with custom DPI and locale | Render chart with French number format using Aspose.Cells and save as PNG
// Tags: set workbook culture Aspose.Cells | chart rendering locale Aspose.Cells | chart to PNG using ImageOrPrintOptions | custom DPI chart export Aspose.Cells | C# chart image generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example demonstrates how to load an Excel workbook, assign a specific CultureInfo to the workbook to control locale‑dependent rendering, create or retrieve a chart, configure ImageOrPrintOptions for single‑page output, and export the chart as a PNG image. It also includes error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a column chart (or retrieve an existing one)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Configure image rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Render the chart on a single page
            };

            // Output image path
            string outputPath = "chart.png";

            // Render the chart to an image file
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart image saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
