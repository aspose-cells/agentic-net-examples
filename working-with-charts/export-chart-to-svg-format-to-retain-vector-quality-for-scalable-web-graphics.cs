// Title: Export the first worksheet chart to an SVG file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook, selects the first chart on the first worksheet, and saves it as an SVG using Aspose.Cells ImageOrPrintOptions. | Show how to configure ImageOrPrintOptions with SaveFormat.Svg for chart export and include checks for missing input files and empty chart collections. | Provide error‑handling logic for file‑not‑found and no‑chart scenarios before calling chart.ToImage to generate the SVG.
// Common Searches: Aspose.Cells C# export chart as SVG vector graphic | How to save an Excel chart to SVG using Aspose.Cells .NET | C# sample code for converting worksheet chart to SVG with ImageOrPrintOptions | Export first chart from workbook to scalable SVG with Aspose.Cells
// Tags: chart.ToImage SVG export Aspose.Cells | ImageOrPrintOptions SaveFormat.Svg C# | export Excel chart to vector graphic | Aspose.Cells chart SVG conversion | C# workbook chart extraction Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering; // Required for ImageOrPrintOptions

// The program loads an Excel workbook, retrieves the first chart from the first worksheet, configures ImageOrPrintOptions with SaveFormat.Svg, and exports the chart to an SVG file while handling missing files and empty chart collections.
class ExportChartToSvg
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.svg";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the chart.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart.
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart.
            Chart chart = worksheet.Charts[0];

            // Configure image options for SVG output.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg
                // Optional: set image dimensions if required.
                // ImageWidth = 800,
                // ImageHeight = 600
            };

            // Export the chart to an SVG file.
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
