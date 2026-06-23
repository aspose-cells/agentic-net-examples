using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Output directory for chart images
            string outputDir = "ChartImages";
            Directory.CreateDirectory(outputDir);

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets
            for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIdx];

                // Iterate through all charts in the current worksheet
                for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
                {
                    Chart chart = sheet.Charts[chartIdx];

                    // Build a unique file name for each chart
                    string chartFileName = Path.Combine(
                        outputDir,
                        $"Chart_Sheet{sheetIdx}_Chart{chartIdx}.png");

                    // Export chart to PNG using the ToImage(string, ImageType) rule
                    chart.ToImage(chartFileName, ImageType.Png);
                }
            }

            Console.WriteLine("All charts have been exported successfully.");
        }
    }
}