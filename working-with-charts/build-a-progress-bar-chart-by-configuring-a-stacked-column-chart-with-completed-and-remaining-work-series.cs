// Title: Create a Progress Bar Chart with a Stacked Column in Aspose.Cells for .NET (C#)
// Description: This example builds an Excel workbook, writes a task name with completed and remaining percentages, adds a stacked column chart, assigns the category axis, creates "Completed" and "Remaining" series, colors them green and light‑gray, removes column gaps, sets an overlap of -100 to merge the series into a single bar, adds a chart title, and saves the file as ProgressBarChart.xlsx.
// Keywords: Aspose.Cells progress bar chart | stacked column chart C# | Aspose.Cells set gap width | Aspose.Cells series overlap | C# Excel progress bar | Aspose.Cells chart series colors | export Aspose.Cells chart as PNG | Aspose.Cells stacked column example
// Common Searches: How to create a progress bar in Excel using Aspose.Cells | Aspose.Cells stacked column chart without gaps | Set overlap -100 in Aspose.Cells column chart | Change series color in Aspose.Cells chart | Export Aspose.Cells chart to image
// Developer Intent: Generate an Excel file that visualizes task completion as a single‑category progress bar using a stacked column chart.
// Use Cases: Project‑management dashboard showing task completion percentages | Status‑report worksheets with printable progress indicators | Automated invoices that display processing stage progress | Production‑line KPI sheets visualizing throughput targets | Learning‑management reports illustrating course completion
// AI Prompts: Add data labels that display the percentage value for each series in the progress bar chart. | Write code to loop through multiple task rows and create a stacked progress bar for each row automatically. | Provide a snippet to export the generated progress bar chart as a PNG file using Aspose.Cells. | Show how to apply conditional coloring so the "Completed" segment turns red when below 50%. | Generate dynamic chart titles that pull the task name from the worksheet cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartDemo
{
    // This example builds an Excel workbook, writes a task name with completed and remaining percentages, adds a stacked column chart, assigns the category axis, creates "Completed" and "Remaining" series, colors them green and light‑gray, removes column gaps, sets an overlap of -100 to merge the series into a single bar, adds a chart title, and saves the file as ProgressBarChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare data for the progress bar (single category)
            // -------------------------------------------------
            // Category label
            sheet.Cells["A2"].PutValue("Task 1");

            // Completed work (e.g., 70%)
            sheet.Cells["B2"].PutValue(70);

            // Remaining work (e.g., 30%)
            sheet.Cells["C2"].PutValue(30);

            // -------------------------------------------------
            // Add a stacked column chart to represent the progress bar
            // -------------------------------------------------
            // Parameters: chart type, top row, left column, bottom row, right column
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A2";

            // Add the "Completed" series
            int completedSeriesIdx = chart.NSeries.Add("B2:B2", true);
            Series completedSeries = chart.NSeries[completedSeriesIdx];
            completedSeries.Name = "Completed";
            completedSeries.Area.ForegroundColor = Color.Green; // Color for completed part

            // Add the "Remaining" series
            int remainingSeriesIdx = chart.NSeries.Add("C2:C2", true);
            Series remainingSeries = chart.NSeries[remainingSeriesIdx];
            remainingSeries.Name = "Remaining";
            remainingSeries.Area.ForegroundColor = Color.LightGray; // Color for remaining part

            // Remove gaps between columns to make it look like a single bar
            chart.GapWidth = 0;          // No space between column clusters
            completedSeries.GapWidth = 0;
            remainingSeries.GapWidth = 0;

            // Optional: set overlap to -100 to ensure the two parts touch each other tightly
            completedSeries.Overlap = -100;
            remainingSeries.Overlap = -100;

            // Add a title to the chart
            chart.Title.Text = "Progress Bar";

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
