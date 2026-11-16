// Title: Add percentage data labels to a visible series in an Excel Progress Bar chart using Aspose.Cells for .NET
// Common Searches: how to display percentage data labels on a progress bar chart with Aspose.Cells C# | Aspose.Cells C# add data labels showing only percentages to chart series | set number format for chart data labels to percent in Aspose.Cells | enable data labels for each series of an Excel chart using Aspose.Cells .NET | add data labels to the first chart in a workbook with Aspose.Cells
// Tags: Aspose.Cells chart data labels percentage | C# Aspose.Cells set data label number format | Excel progress bar chart series data labels | enable data labels visible series Aspose.Cells | Aspose.Cells chart series formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, locates the first chart (a progress bar), iterates through its series, enables data labels to show only the percentage value, formats the labels as whole percentages, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "ProgressBar.xlsx";
            const string outputPath = "ProgressBar_WithDataLabels.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the Progress Bar chart
            Workbook workbook = new Workbook(inputPath);

            // Assume the chart is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the sheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart progressChart = sheet.Charts[0];

            // Add data labels showing exact percentage values to each series
            foreach (Series series in progressChart.NSeries)
            {
                // Enable data labels for the series
                series.DataLabels.ShowPercentage = true;
                series.DataLabels.ShowValue = false;

                // Format the percentage display (e.g., 75%)
                series.DataLabels.NumberFormat = "0%";
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
