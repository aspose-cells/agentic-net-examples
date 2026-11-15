// Title: Bind an Excel column of progress percentages to a stacked column progress bar chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bind column C2:C11 to stacked column chart series | How to set values of a progress bar chart from an Excel range with Aspose.Cells | Create or update a ColumnStacked chart as a progress bar in .NET | Assign worksheet range to chart series programmatically using Aspose.Cells | Load existing workbook and bind progress percentages to chart series in C#
// Tags: bind worksheet range to chart series Aspose.Cells | column stacked progress bar chart .NET | set chart series values from Excel range C# | update existing chart series Aspose.Cells | create progress bar visualization Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, locates or creates a ColumnStacked chart, defines the range C2:C11 containing progress percentages, ensures the chart has a series, assigns the range to the series values, and saves the workbook with the updated progress‑bar visualization.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Try to find an existing ColumnStacked chart (used as a progress‑bar placeholder)
            Chart progressChart = null;
            foreach (Chart c in sheet.Charts)
            {
                if (c.Type == ChartType.ColumnStacked)
                {
                    progressChart = c;
                    break;
                }
            }

            // If not found, create a new ColumnStacked chart
            if (progressChart == null)
            {
                // Add a new chart; Add returns the index of the chart in the collection
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 15, 10);
                progressChart = sheet.Charts[chartIndex];
            }

            // Define the range that holds the progress percentages
            const string progressRange = "C2:C11";

            // Ensure the chart has at least one series
            if (progressChart.NSeries.Count == 0)
            {
                // Add a series; false indicates that data is in columns
                progressChart.NSeries.Add(progressRange, false);
            }

            // Bind the progress percentages to the first series
            progressChart.NSeries[0].Values = progressRange;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
