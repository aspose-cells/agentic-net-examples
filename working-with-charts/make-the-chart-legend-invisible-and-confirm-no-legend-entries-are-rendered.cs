// Title: Hide a chart legend and verify no legend entries using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook with Aspose.Cells, set the first chart's ShowLegend property to false, read the Legend.LegendEntries.Count to ensure it is zero, and save the modified file. | Write C# code that disables a chart's legend, prints the legend visibility flag and entry count, and confirms the legend is invisible with Aspose.Cells.
// Common Searches: Aspose.Cells C# hide chart legend and check entry count | how to disable legend in an Excel chart using Aspose.Cells .NET | verify that a chart legend is not rendered after setting ShowLegend false in Aspose.Cells | C# example for removing legend from first chart in a workbook with Aspose.Cells | Aspose.Cells chart legend visibility test code
// Tags: Aspose.Cells hide chart legend C# | Aspose.Cells ShowLegend false usage | Aspose.Cells chart legend entry count verification | Aspose.Cells remove legend from Excel chart .NET | Aspose.Cells chart visibility check C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads ChartSample.xlsx, accesses the first worksheet and its first chart, sets ShowLegend to false to hide the legend, checks that ShowLegend is false and that the legend entry count is zero, saves the workbook as ChartSample_NoLegend.xlsx, and writes the verification results to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "ChartSample.xlsx";
            const string outputPath = "ChartSample_NoLegend.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Hide the chart legend using the correct API
            chart.ShowLegend = false;

            // Verify legend visibility
            bool legendIsInvisible = !chart.ShowLegend;
            int legendEntryCount = chart.Legend.LegendEntries.Count;

            // Save the modified workbook
            workbook.Save(outputPath);

            // Output verification results
            Console.WriteLine($"Legend visible: {chart.ShowLegend}");
            Console.WriteLine($"Legend entries count: {legendEntryCount}");
            Console.WriteLine($"Legend invisibility confirmed: {legendIsInvisible}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
