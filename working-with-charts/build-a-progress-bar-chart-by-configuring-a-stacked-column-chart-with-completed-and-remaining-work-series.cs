// Title: Generate a stacked column progress bar chart in Excel with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to build a stacked column chart that displays completed and remaining work as a progress bar, applying custom colors to each series. | Adjust the chart configuration to eliminate column gaps, set the chart title, and define the category axis labels programmatically.
// Common Searches: Aspose.Cells C# create stacked column chart for progress bar with completed and remaining percentages | how to set series foreground colors in an Aspose.Cells stacked column chart | remove gaps between columns in Aspose.Cells stacked chart to make solid bars | set category axis data range for stacked column chart using Aspose.Cells | save Excel workbook with progress bar visualization using Aspose.Cells .NET
// Tags: Aspose.Cells stacked column chart progress bar | C# set series foreground color Aspose.Cells | Aspose.Cells remove column gap stacked chart | Aspose.Cells set chart title and category data | Aspose.Cells save workbook as XLSX

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds task data, and constructs a stacked column chart where a green 'Completed' series and a light‑gray 'Remaining' series form a progress‑bar visual. Column gaps are removed for a solid appearance, the chart title and category axis are set, and the workbook is saved as ProgressBarChart.xlsx.
class ProgressBarChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet and name it
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Progress";

            // Header row
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Completed");
            sheet.Cells["C1"].PutValue("Remaining");

            // Sample data rows
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(70);   // 70% completed
            sheet.Cells["C2"].PutValue(30);   // 30% remaining

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(45);   // 45% completed
            sheet.Cells["C3"].PutValue(55);   // 55% remaining

            // Add a stacked column chart (progress bar style)
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIndex];

            // Chart title
            chart.Title.Text = "Progress Bar Chart";

            // Category (X) axis labels – tasks
            chart.NSeries.CategoryData = "A2:A3";

            // Completed series (green)
            int completedSeriesIdx = chart.NSeries.Add("B2:B3", true);
            var completedSeries = chart.NSeries[completedSeriesIdx];
            completedSeries.Name = "Completed";
            completedSeries.Area.ForegroundColor = Color.Green;

            // Remaining series (light gray)
            int remainingSeriesIdx = chart.NSeries.Add("C2:C3", true);
            var remainingSeries = chart.NSeries[remainingSeriesIdx];
            remainingSeries.Name = "Remaining";
            remainingSeries.Area.ForegroundColor = Color.LightGray;

            // Remove gaps between columns to make them look like solid bars
            chart.GapWidth = 0;

            // Save the workbook
            string outputPath = "ProgressBarChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
