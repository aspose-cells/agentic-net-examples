// Title: Disable data labels for the third series in an Excel chart with Aspose.Cells for .NET
// AI Prompts: Set ShowDataLabels = false on the third series of a chart in an existing .xlsx workbook using Aspose.Cells C# API. | Programmatically hide data labels for only the third series while keeping labels on other series in an Excel chart via Aspose.Cells. | Update a workbook so that the chart’s third data series does not display data labels, using Aspose.Cells for .NET.
// Common Searches: aspocells hide data labels third series chart c# | how to turn off data labels for a specific series in an Excel chart using Aspose.Cells | C# Aspose.Cells set ShowDataLabels false for chart series index 2 | remove data labels from only one series in an Excel chart programmatically | Aspose.Cells chart series customization hide labels for selected series
// Tags: Aspose.Cells hide chart series data labels | Excel chart third series ShowDataLabels false C# | Aspose.Cells modify specific chart series property | C# disable data labels for selected series in Excel | Aspose.Cells chart series visibility control

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, verifies a chart with at least three series, disables data labels on the third series by setting ShowDataLabels to false, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least three series
            if (chart.NSeries.Count > 2)
            {
                // Use dynamic to avoid compile‑time dependency on ChartSeries type
                dynamic nSeries = chart.NSeries;
                dynamic thirdSeries = nSeries[2];
                thirdSeries.ShowDataLabels = false;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
