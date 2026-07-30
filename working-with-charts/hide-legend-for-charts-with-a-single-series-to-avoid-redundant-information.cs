// Title: Aspose.Cells for .NET – Hide Chart Legend When Only One Series Exists
// Description: C# example that creates a workbook, adds a column chart, binds a single data series, checks the series count, and disables the legend to avoid redundant labels before saving the file.
// Keywords: Aspose.Cells hide legend | chart legend conditional .NET | single series chart Aspose | ShowLegend false Aspose.Cells | C# Excel chart formatting
// Common Searches: hide legend Aspose.Cells chart single series | Aspose.Cells conditional legend display | C# remove chart legend when only one series | Aspose.Cells ShowLegend property usage
// Developer Intent: Automatically suppress the legend of a chart that contains exactly one data series using Aspose.Cells for .NET.
// Use Cases: Generate a column chart from a data range and automatically hide the legend when only one series is plotted. | Process multiple charts in a workbook, applying the same rule to remove legends from single‑series charts. | Create clean reporting workbooks where single‑series charts are displayed without unnecessary legend entries.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart and hide its legend only if the chart has a single series. | Show how to iterate over all charts in a workbook and set ShowLegend = false for charts with exactly one series. | Explain the steps to check a chart's NSeries count in Aspose.Cells and control legend visibility programmatically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a column chart, binds a single data series, checks the series count, and disables the legend to avoid redundant labels before saving the file.
    public class HideLegendForSingleSeriesChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Hide legend if only one series exists
                if (chart.NSeries.Count == 1)
                {
                    chart.ShowLegend = false;
                }

                // Determine output file path
                string outputFile = "HideLegendSingleSeriesChart.xlsx";

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputFile)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideLegendForSingleSeriesChart.Run();
        }
    }
}
