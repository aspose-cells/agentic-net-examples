// Title: How to set a custom foreground color for the third series of a chart in an existing XLSX workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change color of third chart series in an existing XLSX file | set specific series color in a chart with Aspose.Cells .NET | modify chart series foreground color programmatically using Aspose.Cells | C# example to update chart series color and save workbook with Aspose.Cells | find chart with at least three series and change third series color Aspose.Cells
// Tags: set chart series foreground color Aspose.Cells | modify third series color XLSX Aspose.Cells | access NSeries collection .NET | save modified workbook Aspose.Cells | chart series color customization Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing XLSX workbook, locates the first chart containing at least three series, accesses the third series (index 2), sets its area foreground color to red, ensures the output directory exists, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Find a chart that contains at least three series
            Chart? targetChart = null;
            foreach (Chart chart in sheet.Charts)
            {
                if (chart.NSeries.Count >= 3)
                {
                    targetChart = chart;
                    break;
                }
            }

            if (targetChart != null)
            {
                // Access the third series (zero‑based index 2)
                var thirdSeries = targetChart.NSeries[2];

                // Example: set an explicit color for the third series
                thirdSeries.Area.ForegroundColor = Color.FromArgb(255, 0, 0);
            }
            else
            {
                Console.WriteLine("No chart with at least three series was found.");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
