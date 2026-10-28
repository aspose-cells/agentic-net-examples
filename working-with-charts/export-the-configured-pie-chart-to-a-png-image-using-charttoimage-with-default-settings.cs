// Title: Export a pie chart from an Excel workbook to a PNG file using Aspose.Cells Chart.ToImage in C#
// Common Searches: C# Aspose.Cells export first worksheet chart as PNG | How to save an Excel pie chart as an image with Aspose.Cells | Using Chart.ToImage to generate PNG from an Excel chart in .NET | Aspose.Cells convert Excel chart to PNG file with default settings | Export Excel chart to PNG without specifying image options in C#
// Tags: Aspose.Cells chart to PNG conversion | Chart.ToImage method usage | C# export Excel pie chart image | default settings image export Aspose.Cells | save worksheet chart as PNG file

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing workbook, checks for charts on the first worksheet, retrieves the first chart (assumed to be a pie chart), and exports it to a PNG file using the Chart.ToImage method with default options, while handling missing files and runtime exceptions.
class ExportChartToPng
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "pieChart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the configured pie chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts were found in the worksheet.");
                return;
            }

            // Retrieve the first chart (assumed to be the pie chart)
            Chart chart = sheet.Charts[0];

            // Export the chart to a PNG image using default settings
            chart.ToImage(outputPath);

            Console.WriteLine($"Pie chart exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
