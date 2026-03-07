using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ConvertChartToSvg
{
    static void Main()
    {
        // Load the workbook that contains the chart
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Retrieve the first chart
        Chart chart = worksheet.Charts[0];

        // Set up SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions
        {
            ImageType = ImageType.Svg,   // Specify SVG output format
            FitToViewPort = true,        // Optional: fit SVG to viewport
            CssPrefix = "chart-"         // Optional: CSS class prefix
        };

        // Export the chart to an SVG file
        chart.ToImage("output_chart.svg", svgOptions);

        Console.WriteLine("Chart successfully exported to SVG.");
    }
}