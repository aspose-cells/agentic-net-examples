using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ChartToImageExample
{
    static void Main()
    {
        // Load the workbook containing the chart
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No chart found in the worksheet.");
            return;
        }

        // Get the first chart
        Chart chart = worksheet.Charts[0];

        // Configure image rendering options for high fidelity
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,          // Output format
            HorizontalResolution = 300,         // DPI X
            VerticalResolution = 300            // DPI Y
        };

        // Save the chart as an image using the provided ToImage method
        string outputImagePath = "chart_image.png";
        chart.ToImage(outputImagePath, options);

        Console.WriteLine($"Chart image successfully saved to: {outputImagePath}");
    }
}