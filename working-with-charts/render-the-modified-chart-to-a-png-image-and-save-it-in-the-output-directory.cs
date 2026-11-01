// Title: Export a modified Excel chart to PNG using Aspose.Cells for .NET (C# example)
// AI Prompts: Write C# code that opens an existing .xlsx workbook, changes the first chart's title, adds a new series from a range, switches the chart type to Column, and saves the chart as a PNG file with Aspose.Cells. | Show how to create the output directory if it doesn't exist and use Aspose.Cells Chart.ToImage to render a modified chart to a PNG image in a .NET application. | Demonstrate exporting a specific worksheet chart to an image while applying title and series modifications using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to save a chart as PNG after modifying it | Aspose.Cells export chart image from workbook with custom title | Render Excel chart to PNG file using Aspose.Cells Chart.ToImage method | Create output folder and export modified chart to PNG in .NET
// Tags: Aspose.Cells chart image export | update chart title C# Aspose.Cells | add data series to Excel chart Aspose.Cells | Chart.ToImage method Aspose.Cells | create output folder C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, verifies a chart exists on the first worksheet, updates the chart title, adds a new data series, changes the chart type to Column, ensures the output directory is present, and then renders the modified chart directly to a PNG file using Aspose.Cells' Chart.ToImage method.
class ChartToPngExample
{
    static void Main()
    {
        try
        {
            // Input workbook path.
            string inputPath = "input.xlsx";

            // Verify the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the chart.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0).
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart.
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart in the worksheet.
            Chart chart = sheet.Charts[0];

            // ----- Begin modifications to the chart (example) -----
            // Change the chart title.
            chart.Title.Text = "Modified Chart Title";

            // Add a new series (example: using data from A1:B5).
            int seriesIndex = chart.NSeries.Add("A1:B5", true);
            chart.NSeries[seriesIndex].Name = "New Series";

            // Set the chart type (example: Column).
            chart.Type = ChartType.Column;
            // ----- End of chart modifications -----

            // Define output path and ensure directory exists.
            string outputPath = Path.Combine("output", "chart.png");
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            Directory.CreateDirectory(outputDir);

            // Render the chart directly to a PNG file.
            chart.ToImage(outputPath, ImageType.Png);

            Console.WriteLine($"Chart rendered and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
