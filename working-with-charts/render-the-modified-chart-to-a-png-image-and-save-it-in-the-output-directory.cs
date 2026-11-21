// Title: Export a modified Excel chart to PNG using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export chart as PNG after changing title | How to save an Excel chart to a PNG file with Aspose.Cells in .NET | Aspose.Cells chart.ToImage example for PNG output in C# | Render first worksheet chart to PNG using ImageOrPrintOptions in Aspose.Cells | Convert Excel chart to PNG image programmatically with Aspose.Cells .NET
// Tags: Aspose.Cells chart PNG export C# | ImageOrPrintOptions OnePagePerSheet usage | modify chart title Aspose.Cells | Aspose.Cells chart image rendering | save Excel chart as PNG .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads a workbook, accesses the first worksheet's first chart, updates its title, configures ImageOrPrintOptions, and uses chart.ToImage to render the chart as a PNG file saved in the output directory, with error handling for missing files and absent charts.
class ChartToPng
{
    static void Main()
    {
        // Define input and output directories
        string inputDir = @"./input/";
        string outputDir = @"./output/";

        try
        {
            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);

            // Build full path to the input workbook
            string workbookPath = Path.Combine(inputDir, "sample.xlsx");

            // Verify that the input file exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Input file not found: {workbookPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet and its first chart
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            Chart chart = worksheet.Charts[0];

            // Modify the chart as needed (example: change the title)
            chart.Title.Text = "Modified Chart";

            // Set image rendering options (format is inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the chart to a PNG file
            string outputPath = Path.Combine(outputDir, "chart.png");
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart rendered and saved as PNG: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
