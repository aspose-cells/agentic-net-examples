// Title: Export a pie chart from an Excel workbook to a PNG file using Aspose.Cells Chart.ToImage in C#
// AI Prompts: Write C# code that opens an Excel file, selects the first chart (assumed to be a pie chart), and saves it as a PNG using Aspose.Cells Chart.ToImage with default ImageOrPrintOptions. | Show how to use ImageOrPrintOptions together with Chart.ToImage to render any Excel chart to a PNG image in a .NET application.
// Common Searches: c# aspose.cells export first chart in workbook to png | how to render an Excel pie chart as a PNG using Aspose.Cells Chart.ToImage | asp.net save Excel chart to image with default settings Aspose.Cells
// Tags: Aspose.Cells Chart.ToImage PNG export | C# export Excel chart to image | ImageOrPrintOptions default configuration | render pie chart as PNG Aspose.Cells | export workbook chart to PNG .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example loads 'input.xlsx', accesses the first worksheet, retrieves the first chart (expected to be a pie chart), and uses Aspose.Cells' Chart.ToImage method with default ImageOrPrintOptions to generate 'pieChart.png'. It includes checks for missing files and absent charts, and reports success or errors to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "pieChart.png";

            // Verify the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart (assumed to be a pie chart)
            Chart pieChart = sheet.Charts[0];

            // Set image export options (default format is PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions();

            // Export the chart to a PNG image
            pieChart.ToImage(outputPath, options);
            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
