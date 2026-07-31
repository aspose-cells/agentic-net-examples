// Title: Aspose.Cells .NET – Clear Existing Chart Trendlines and Add New Linear & Exponential Trendlines (C#)
// Description: Shows how to build a workbook, create a line chart, remove all trendlines from each series with TrendLines.Clear(), then insert fresh linear and exponential trendlines (including equation, R‑squared and custom colors) and save the file.
// Keywords: Aspose.Cells C# clear chart trendlines | remove trendlines Aspose.Cells .NET | add linear trendline Aspose.Cells | add exponential trendline Aspose.Cells | TrendLines.Clear() | Aspose.Cells chart manipulation | programmatic trendline management | Excel chart trendline C# | dynamic chart trendlines Aspose | Aspose.Cells chart template
// Common Searches: how to delete trendlines from an Aspose.Cells chart | Aspose.Cells clear all trendlines before adding new ones | C# remove chart trendlines Aspose.Cells | reset trendlines in Aspose.Cells line chart | add linear and exponential trendlines with Aspose.Cells
// Developer Intent: Remove any existing trendlines from a chart’s series before inserting new analytical trendlines.
// Use Cases: Regenerate a report chart where previous trendlines must be cleared and updated linear/exponential trendlines applied to the latest data. | Create a reusable workbook template that automatically wipes old trendlines each time it is generated, ensuring only the required analytical lines appear. | Build an interactive dashboard that resets trendlines on user‑driven data changes, then adds custom polynomial or moving‑average trendlines as needed.
// AI Prompts: Write C# code that iterates through all series in an Aspose.Cells chart and calls TrendLines.Clear() to delete existing trendlines. | Provide an example of adding a polynomial trendline of order 3 after clearing previous trendlines in an Aspose.Cells line chart. | Explain how to preserve custom formatting (color, equation display, R‑squared) when recreating trendlines after they have been cleared in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsTrendlineManagement
{
    // Shows how to build a workbook, create a line chart, remove all trendlines from each series with TrendLines.Clear(), then insert fresh linear and exponential trendlines (including equation, R‑squared and custom colors) and save the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);          // X values: 1,2,3,4,5
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y values: 10,20,30,40,50
                }

                // Add a line chart
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series
                chart.NSeries.Add("B2:B6", true);
                // Assign category (X) data to the series
                chart.NSeries[0].XValues = "A2:A6";

                // ------------------------------------------------------------
                // Remove any existing trendlines from all series in the chart
                // ------------------------------------------------------------
                foreach (Series series in chart.NSeries)
                {
                    series.TrendLines.Clear();
                }

                // ------------------------------------------------------------
                // Add new analytical trendlines as required
                // ------------------------------------------------------------
                // Example: add a linear trendline
                int linearTrendIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
                Trendline linearTrend = chart.NSeries[0].TrendLines[linearTrendIndex];
                linearTrend.Name = "Linear Trend";
                linearTrend.DisplayEquation = true;
                linearTrend.DisplayRSquared = true;
                linearTrend.Color = Color.Blue;

                // Example: add an exponential trendline
                int expTrendIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
                Trendline expTrend = chart.NSeries[0].TrendLines[expTrendIndex];
                expTrend.Name = "Exponential Trend";
                expTrend.DisplayEquation = true;
                expTrend.DisplayRSquared = true;
                expTrend.Color = Color.Red;

                // Save the workbook to a file
                string outputPath = "ChartWithCleanedTrendlines.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
