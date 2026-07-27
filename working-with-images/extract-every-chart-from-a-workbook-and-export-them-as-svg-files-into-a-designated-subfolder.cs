using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ExportChartsToSvg
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Folder where SVG files will be saved
        string outputFolder = "ChartsSvg";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all worksheets
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIdx];

            // Iterate through all charts in the worksheet
            for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
            {
                Chart chart = sheet.Charts[chartIdx];

                // Build a unique file name for each chart
                string fileName = $"Chart_Sheet{sheetIdx}_Index{chartIdx}.svg";
                string filePath = Path.Combine(outputFolder, fileName);

                // Export chart to SVG using the ToImage method with ImageType.Svg
                chart.ToImage(filePath, ImageType.Svg);
            }
        }

        Console.WriteLine("All charts have been exported to SVG files.");
    }
}