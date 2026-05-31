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
            string workbookPath = "input.xlsx";

            // Load the workbook (lifecycle rule: use load)
            Workbook workbook = new Workbook(workbookPath);

            // Define the output folder for SVG files and ensure it exists
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "ExportedCharts");
            Directory.CreateDirectory(outputFolder);

            // Iterate through all worksheets
            for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIdx];
                ChartCollection charts = sheet.Charts;

                // Iterate through all charts in the current worksheet
                for (int chartIdx = 0; chartIdx < charts.Count; chartIdx++)
                {
                    Chart chart = charts[chartIdx];

                    // Build a unique file name for each chart
                    string fileName = $"Sheet{sheetIdx + 1}_Chart{chartIdx + 1}.svg";
                    string filePath = Path.Combine(outputFolder, fileName);

                    // Export the chart to SVG using the ToImage method with ImageType.Svg
                    chart.ToImage(filePath, ImageType.Svg);
                }
            }

            Console.WriteLine($"All charts have been exported to SVG files in folder: {outputFolder}");
        }
    }
}