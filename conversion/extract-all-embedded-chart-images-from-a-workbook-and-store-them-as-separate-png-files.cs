using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartImageExtractor
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Directory where extracted chart images will be saved
            string outputDir = "ChartImages";
            Directory.CreateDirectory(outputDir);

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];

                // Iterate through all charts in the current worksheet
                for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
                {
                    Chart chart = sheet.Charts[chartIndex];

                    // Build a unique file name for each chart image
                    string imageFileName = $"Sheet{wsIndex + 1}_Chart{chartIndex + 1}.png";
                    string imagePath = Path.Combine(outputDir, imageFileName);

                    // Save the chart as PNG using the ToImage(string, ImageType) method (rule)
                    chart.ToImage(imagePath, ImageType.Png);
                }
            }

            Console.WriteLine("All chart images have been extracted to: " + Path.GetFullPath(outputDir));
        }
    }
}