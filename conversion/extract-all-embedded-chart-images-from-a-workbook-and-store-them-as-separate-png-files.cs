// Title: Extract embedded Excel chart images to PNG files with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each chart in an Excel workbook to separate PNG files | How to save Excel chart objects as images using Aspose.Cells .NET | Iterate through worksheets and extract chart pictures to PNG with Aspose.Cells | Aspose.Cells chart.ToImage method example for batch chart extraction | Extract all chart images from .xlsx using Aspose.Cells in C#
// Tags: Aspose.Cells chart.ToImage PNG export | multiple chart image export from Excel | C# iterate worksheets for chart conversion | save workbook charts as separate PNG files | chart image export API Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing .xlsx workbook, walks through every worksheet and each chart on those sheets, and uses the Chart.ToImage method to write each chart to a uniquely named PNG file (Chart_0.png, Chart_1.png, …). It handles missing files and export errors, and reports the total number of extracted chart images.
class ExtractChartImages
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            int chartIndex = 0; // Counter for naming output files

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Build a unique file name for the chart image
                        string fileName = $"Chart_{chartIndex}.png";

                        // Export the chart directly to a PNG file
                        chart.ToImage(fileName);

                        chartIndex++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"Extracted {chartIndex} chart image(s).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
