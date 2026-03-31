using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class RenderPieChartPreview
{
    static void Main()
    {
        // Load the workbook that contains the pie chart
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Retrieve the first chart (assumed to be a pie chart)
        Chart chart = worksheet.Charts[0];

        // Define the output image file path
        string outputImagePath = "pie_chart_preview.png";

        // Render the chart to an image file (PNG format)
        chart.ToImage(outputImagePath, ImageType.Png);

        Console.WriteLine($"Chart image successfully saved to: {outputImagePath}");
    }
}