using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ChartToSvg
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart (you can select a different chart by index if needed)
        Chart chart = worksheet.Charts[0];

        // Configure SVG rendering options
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.ImageType = ImageType.Svg;          // Explicitly set output format to SVG
        svgOptions.FitToViewPort = true;              // Make SVG fit the viewport for better scaling
        svgOptions.EmbeddedFontType = SvgEmbeddedFontType.Woff; // Embed fonts to preserve appearance
        svgOptions.CssPrefix = "chart";               // Optional CSS prefix to avoid naming conflicts

        // Export the chart to an SVG file while preserving visual fidelity
        chart.ToImage("chart_output.svg", svgOptions);

        Console.WriteLine("Chart has been successfully exported to SVG.");
    }
}