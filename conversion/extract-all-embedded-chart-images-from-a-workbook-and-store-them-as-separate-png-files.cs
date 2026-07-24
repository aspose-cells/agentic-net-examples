// Title: C# – Extract All Embedded Excel Charts as PNG Images with Aspose.Cells
// Description: Load an Excel workbook using Aspose.Cells, loop through each worksheet and its charts, and export every chart to a separate PNG file via the Chart.ToImage method.
// Keywords: Aspose.Cells chart export C# | save Excel chart as PNG | extract embedded charts .NET | Chart.ToImage example | batch export Excel charts
// Common Searches: how to export all charts from an Excel file to PNG using Aspose.Cells | C# code to save each worksheet chart as an image | Aspose.Cells extract chart images programmatically
// Developer Intent: Programmatically retrieve every chart in a workbook and write each one to an individual PNG file.
// Use Cases: Create thumbnail previews of workbook charts for a web portal. | Generate image assets for documentation, reports, or presentations. | Automate batch conversion of Excel charts to PNG for archival purposes.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through all worksheets in a workbook and save each chart as a PNG image. | Provide a reusable method that accepts a workbook path and an output folder, extracts all chart images, and returns the list of saved file paths. | Explain how to adjust PNG export settings such as resolution and background color when using Chart.ToImage in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartImageExtractor
{
    // Load an Excel workbook using Aspose.Cells, loop through each worksheet and its charts, and export every chart to a separate PNG file via the Chart.ToImage method.
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path (replace with your actual file)
            string workbookPath = "input.xlsx";

            // Output directory for extracted chart images
            string outputDir = "ExtractedCharts";
            Directory.CreateDirectory(outputDir);

            // Load the workbook (using Aspose.Cells load rule)
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
                    string chartFileName = $"{sheet.Name}_Chart{chartIndex}.png";
                    string chartFilePath = Path.Combine(outputDir, chartFileName);

                    // Save the chart as PNG (using Chart.ToImage(string, ImageType) rule)
                    chart.ToImage(chartFilePath, ImageType.Png);

                    Console.WriteLine($"Saved chart image: {chartFilePath}");
                }
            }

            Console.WriteLine("All chart images have been extracted.");
        }
    }
}
