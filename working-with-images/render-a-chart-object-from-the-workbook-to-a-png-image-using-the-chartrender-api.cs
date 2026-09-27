// Title: Export the first chart from an Excel worksheet to a PNG image using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx workbook, retrieves the first chart on the first worksheet, and saves it as a PNG file using Aspose.Cells Chart.ToImage. | Show how to add error handling for missing workbook files or absent charts when exporting an Excel chart to PNG with Aspose.Cells in a .NET console application.
// Common Searches: asp.net console export excel chart to png using aspose.cells | c# example for Chart.ToImage method Aspose.Cells | how to save an Excel chart as an image file with Aspose.Cells library | render first worksheet chart to png without opening Excel | Aspose.Cells ChartRender API usage in C#
// Tags: Aspose.Cells chart export to PNG | Chart.ToImage method C# | C# render Excel chart as image file | save first worksheet chart as PNG | error handling missing chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook, checks for charts on the first worksheet, extracts the first chart, and uses the Chart.ToImage method to render and save it as a PNG image, with safeguards for missing files and absent charts.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputImagePath = "chart.png";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Render the chart directly to a PNG image file (default format is PNG)
            chart.ToImage(outputImagePath);

            Console.WriteLine($"Chart image saved to: {outputImagePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
