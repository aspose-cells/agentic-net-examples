// Title: Create a stacked column progress bar chart in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to generate a progress bar chart with stacked columns using Aspose.Cells C# | Aspose.Cells C# set custom fill colors for stacked column series | hide legend in an Aspose.Cells chart programmatically | add task names as category axis labels in an Aspose.Cells stacked column chart | save Excel file with a progress bar visualization using Aspose.Cells .NET
// Tags: custom series fill colors Aspose.Cells | chart legend visibility Aspose.Cells | category axis task labels Aspose.Cells | export progress bar chart to Excel Aspose.Cells | stacked column progress visualization Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartExample
{
    // The example creates a new workbook, fills it with task names and completed/remaining percentages, adds a stacked column chart that acts as a progress bar, customizes the series colors (green for completed, light gray for remaining), assigns task names to the category axis, hides the legend, and saves the workbook as ProgressBarChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for progress bar chart
                // Column A: Task names
                // Column B: Completed work (percentage)
                // Column C: Remaining work (percentage)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Completed");
                sheet.Cells["C1"].PutValue("Remaining");

                // Example rows
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(70);   // 70% completed
                sheet.Cells["C2"].PutValue(30);   // 30% remaining

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["C3"].PutValue(55);

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(90);
                sheet.Cells["C4"].PutValue(10);

                // Add a stacked column chart (acts as a progress bar)
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Progress Bar Chart";

                // Add Completed series (first series)
                int completedSeriesIndex = chart.NSeries.Add("B2:B4", true);
                Series completedSeries = chart.NSeries[completedSeriesIndex];
                completedSeries.Name = "Completed";
                // Set fill color for completed portion (e.g., green)
                completedSeries.Area.ForegroundColor = Color.Green;

                // Add Remaining series (second series)
                int remainingSeriesIndex = chart.NSeries.Add("C2:C4", true);
                Series remainingSeries = chart.NSeries[remainingSeriesIndex];
                remainingSeries.Name = "Remaining";
                // Set fill color for remaining portion (e.g., light gray)
                remainingSeries.Area.ForegroundColor = Color.LightGray;

                // Set category axis labels (Task names)
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: Hide the legend if not needed
                chart.ShowLegend = false;

                // Save the workbook to a file in the current directory
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ProgressBarChart.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for debugging
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
