using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ExportChartsToSvg
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIdx];

                // Iterate through all charts in the current worksheet
                for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
                {
                    Chart chart = sheet.Charts[chartIdx];

                    // Build a unique file name for each chart
                    string svgFileName = $"Chart_Sheet{sheetIdx}_Chart{chartIdx}.svg";

                    // Export the chart to SVG using the ToImage method with ImageType.Svg
                    chart.ToImage(svgFileName, ImageType.Svg);
                }
            }

            Console.WriteLine("All charts have been exported as SVG files.");
        }
    }
}