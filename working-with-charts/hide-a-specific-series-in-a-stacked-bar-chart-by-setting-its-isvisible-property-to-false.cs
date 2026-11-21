// Title: Remove a specific series from a stacked bar chart in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove second series from stacked bar chart | how to delete a series from an Excel chart with Aspose.Cells .NET | programmatically hide a series in a stacked bar chart using Aspose.Cells | C# Aspose.Cells chart NSeries.RemoveAt example | remove chart series from workbook using Aspose.Cells API
// Tags: Aspose.Cells remove chart series | C# stacked bar chart series deletion | Excel chart NSeries.RemoveAt Aspose.Cells | .NET hide series from Excel chart | Aspose.Cells chart series visibility control

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook, accesses the first worksheet and its first stacked bar chart, then removes the series at a given zero‑based index from the chart's NSeries collection using Aspose.Cells for .NET, and finally saves the modified workbook.
class HideChartSeries
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a stacked bar chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = sheet.Charts[0];

            // Verify the chart type is a stacked bar chart (BarStacked)
            if (chart.Type != ChartType.BarStacked)
            {
                Console.WriteLine("The chart is not a stacked bar chart.");
                return;
            }

            // Specify the series index you want to hide (0‑based)
            int seriesIndexToHide = 1; // example: hide the second series

            // Verify the series index is within range
            if (seriesIndexToHide < 0 || seriesIndexToHide >= chart.NSeries.Count)
            {
                Console.WriteLine("Series index out of range.");
                return;
            }

            // Hide the specified series by removing it from the chart's series collection
            chart.NSeries.RemoveAt(seriesIndexToHide);

            // Save the workbook with the modified chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
