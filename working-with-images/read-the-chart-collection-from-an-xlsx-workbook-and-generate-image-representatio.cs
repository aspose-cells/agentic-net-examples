using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ExportChartsToImages
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Directory where chart images will be saved
        string outputDir = "ChartImages";
        Directory.CreateDirectory(outputDir);

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all worksheets
        for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet sheet = workbook.Worksheets[wsIndex];
            ChartCollection charts = sheet.Charts;

            // Iterate through all charts in the current worksheet
            for (int chartIndex = 0; chartIndex < charts.Count; chartIndex++)
            {
                Chart chart = charts[chartIndex];

                // Build a unique file name for each chart image
                // Example: Sheet1_Chart0.png
                string sheetName = sheet.Name;
                string imageFileName = $"{sheetName}_Chart{chartIndex}.png";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Export the chart to an image file (chart.ToImage rule)
                // The file extension determines the image format (PNG in this case)
                chart.ToImage(imagePath);

                Console.WriteLine($"Exported chart [{chartIndex}] from worksheet \"{sheetName}\" to \"{imagePath}\"");
            }
        }

        // Optionally, save the workbook after processing (save rule)
        // This step is not required for chart export but demonstrates the save rule usage
        string savedWorkbookPath = Path.Combine(outputDir, "processed_workbook.xlsx");
        workbook.Save(savedWorkbookPath);
        Console.WriteLine($"Workbook saved to \"{savedWorkbookPath}\"");
    }
}