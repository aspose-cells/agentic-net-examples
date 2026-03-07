using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartToSvg
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a chart
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Configure SVG rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg
            };

            // Export the chart to an SVG file
            chart.ToImage("chart_output.svg", imgOptions);

            Console.WriteLine("Chart exported to SVG successfully.");
        }
    }
}