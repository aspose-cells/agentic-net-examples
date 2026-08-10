// Title: C# – Batch Export All Excel Charts to PNG Files Using Aspose.Cells
// Description: Load an Excel workbook with Aspose.Cells, iterate through every worksheet and chart, and save each chart as an individual PNG image via Chart.ToImage. The routine creates a dedicated output folder and generates unique file names based on sheet and chart indices.
// Keywords: Aspose.Cells C# export chart PNG | batch chart image extraction .NET | save Excel charts as PNG files | export all workbook charts | Chart.ToImage Aspose.Cells
// Common Searches: export every chart in Excel to PNG C# | Aspose.Cells batch chart export example | how to save Excel charts as images programmatically | C# loop worksheets export charts Aspose
// Developer Intent: Automatically generate separate PNG files for each chart contained in an Excel workbook.
// Use Cases: Create image assets for reporting dashboards by extracting all charts from a template workbook. | Build a web gallery of chart thumbnails through bulk PNG export. | Automate documentation workflows that require chart images for PDFs or presentations.
// AI Prompts: Generate a reusable method that accepts a workbook path and output directory, then exports all charts to PNG using Aspose.Cells. | Enhance the script to prepend chart titles to the PNG filenames while handling duplicate titles. | Add robust error handling to log worksheets without charts and continue processing the remaining sheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Load an Excel workbook with Aspose.Cells, iterate through every worksheet and chart, and save each chart as an individual PNG image via Chart.ToImage. The routine creates a dedicated output folder and generates unique file names based on sheet and chart indices.
class ExportChartsToPng
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Directory where chart images will be saved
        string outputDir = "ChartsOutput";
        Directory.CreateDirectory(outputDir);

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Loop through each worksheet
        for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet sheet = workbook.Worksheets[wsIndex];

            // Loop through each chart in the worksheet
            for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
            {
                Chart chart = sheet.Charts[chartIndex];

                // Create a unique file name for the chart image
                string chartFile = Path.Combine(outputDir,
                    $"Sheet{wsIndex}_Chart{chartIndex}.png");

                // Export the chart to a PNG file
                chart.ToImage(chartFile, ImageType.Png);

                Console.WriteLine($"Chart exported: {chartFile}");
            }
        }
    }
}
