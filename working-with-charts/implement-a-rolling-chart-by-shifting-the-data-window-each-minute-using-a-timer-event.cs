// Title: Create a rolling line chart that updates every minute with Aspose.Cells for .NET
// Description: This example builds a workbook with 100 DateTime‑numeric rows, adds a line chart that shows a 10‑row moving window, and uses System.Timers.Timer to shift the data range each minute, recalculate the chart, and save the file.
// Keywords: Aspose.Cells rolling chart | C# timer chart update | dynamic Excel chart .NET | moving window line chart | real‑time chart refresh Aspose
// Common Searches: Aspose.Cells update chart series on a timer | C# rolling chart every minute | how to shift Excel chart data window programmatically | dynamic line chart with Aspose.Cells | timer based chart refresh in .NET
// Developer Intent: Implement a line chart that automatically moves its data window at one‑minute intervals.
// Use Cases: Live sensor dashboard that scrolls forward as new readings arrive. | Performance monitor that continuously displays the latest N data points. | Time‑series log viewer that auto‑advances without manual refresh.
// AI Prompts: Show how to change the timer interval to 30 seconds while keeping the rolling chart functional. | Explain how to store the current startRow in the workbook so the chart resumes at the same position after reopening. | Provide code to add a second data series and synchronize both series during each timer tick.

using System;
using System.IO;
using System.Timers;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example builds a workbook with 100 DateTime‑numeric rows, adds a line chart that shows a 10‑row moving window, and uses System.Timers.Timer to shift the data range each minute, recalculate the chart, and save the file.
class RollingChartDemo
{
    static void Main()
    {
        try
        {
            // ---------- Create workbook and fill sample time‑based data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 100 rows of data: column A = DateTime, column B = numeric value
            for (int i = 0; i < 100; i++)
            {
                sheet.Cells[i, 0].PutValue(DateTime.Now.AddMinutes(i));
                sheet.Cells[i, 1].PutValue(i);
            }

            // ---------- Add a line chart that will display a moving window ----------
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            int windowSize = 10;   // number of rows shown in the chart at any time
            int startRow = 0;      // zero‑based index of the first row in the window

            // Initial data range for the series (Aspose.Cells uses 1‑based cell references)
            chart.NSeries.Add($"B{startRow + 1}:B{startRow + windowSize}", true);
            Series series = chart.NSeries[0] as Series;
            if (series != null)
            {
                // Set category (X) values
                series.XValues = $"A{startRow + 1}:A{startRow + windowSize}";
            }
            chart.Title.Text = "Rolling Chart (updates every minute)";

            // ---------- Timer that shifts the data window every minute ----------
            System.Timers.Timer timer = new System.Timers.Timer(60_000); // 60,000 ms = 1 minute
            timer.Elapsed += (sender, e) =>
            {
                try
                {
                    // Move the window one row down; wrap around when reaching the end
                    startRow++;
                    if (startRow + windowSize > 100)
                        startRow = 0;

                    // Update the series data range to the new window
                    chart.NSeries[0].Values = $"B{startRow + 1}:B{startRow + windowSize}";
                    Series s = chart.NSeries[0] as Series;
                    if (s != null)
                    {
                        s.XValues = $"A{startRow + 1}:A{startRow + windowSize}";
                    }

                    // Re‑calculate the chart so the changes are reflected
                    chart.Calculate();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Timer error: {ex.Message}");
                }
            };
            timer.Start();

            // Keep the console app alive until the user decides to stop it
            Console.WriteLine("Rolling chart is updating every minute. Press ENTER to stop and save.");
            Console.ReadLine();

            timer.Stop();

            // ---------- Save the workbook ----------
            string outputPath = "RollingChart.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
