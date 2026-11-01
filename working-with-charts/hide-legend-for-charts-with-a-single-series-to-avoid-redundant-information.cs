// Title: Hide the legend of a single‑series column chart in Excel using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and automatically disables its legend when the chart contains only one series. | Demonstrate how to use Chart.ShowLegend together with NSeries.Count to conditionally hide a legend in an Excel chart via Aspose.Cells. | Provide a complete example that builds a workbook, adds a single‑series column chart, suppresses the redundant legend, and saves the file.
// Common Searches: Aspose.Cells hide chart legend when only one series C# | conditional legend visibility for Excel charts using Aspose.Cells .NET | how to check NSeries count and remove legend in Aspose.Cells | create column chart without legend for single series in Aspose.Cells | Chart.ShowLegend false example Aspose.Cells C#
// Tags: Aspose.Cells conditional chart legend visibility | single-series chart hide legend .NET | Chart.ShowLegend property Aspose.Cells | Excel column chart without legend Aspose.Cells | detect NSeries count Aspose.Cells chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, fills month‑sales data, adds a column chart, checks if the chart has only one series, sets Chart.ShowLegend to false to remove the redundant legend, and saves the workbook as ChartWithConditionalLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a single‑series chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(250);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Define the series and category data ranges
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend if the chart contains only one series
            if (chart.NSeries.Count == 1)
            {
                // Use Chart.ShowLegend to control visibility
                chart.ShowLegend = false;
            }

            // Determine output file path
            string outputPath = "ChartWithConditionalLegend.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
