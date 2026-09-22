// Title: Export an Excel chart to PNG while preserving data label visibility with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx workbook, activates value and category name data labels for every series in the first chart, and saves the chart as a PNG file using Aspose.Cells. | Demonstrate how to configure image rendering options and invoke the chart rendering method to produce a PNG image that retains all data labels.
// Common Searches: Aspose.Cells C# export chart to PNG with data labels shown | How to keep series values visible when converting Excel chart to image | Chart image export preserving category names using Aspose.Cells | C# render Excel chart as PNG while displaying data labels | ImageOrPrintOptions settings for chart-to-image conversion in Aspose.Cells
// Tags: export chart to PNG Aspose.Cells | enable data labels before chart image rendering | chart image rendering options C# | preserve series values in chart PNG export | ImageOrPrintOptions OnePagePerSheet usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, turns on value and category name data labels for each series in the first chart, and exports the chart as a PNG image using Aspose.Cells with appropriate rendering options.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Enable data labels for each series and show values
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;
                series.DataLabels.ShowCategoryName = true;
            }

            // Set image options for PNG export (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the chart directly to an image file
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart image saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
