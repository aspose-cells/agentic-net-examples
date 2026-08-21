// Title: Aspose.Cells for .NET – Show percentage labels on the visible series of a stacked‑bar progress chart
// Description: Creates an XLSX workbook, adds task data, builds a stacked bar chart used as a progress bar, makes the "Remaining" series transparent, disables its labels, and configures the "Completed" series to display only the percentage (0 % format) inside each bar.
// Keywords: Aspose.Cells | .NET | C# | stacked bar chart | progress bar chart | percentage data labels | hide series | transparent series | inside end label | number format 0% | chart customization | Excel automation
// Common Searches: Aspose.Cells display percentage on stacked bar | C# progress bar chart data label | hide series in Aspose.Cells chart | set label position inside end Aspose.Cells | format chart data label as percent .NET
// Developer Intent: Add percentage data labels to the visible (completed) series of a stacked‑bar progress chart while hiding the remaining series.
// Use Cases: Generate a progress‑bar style chart where only the completed portion is visible and each bar shows its completion percent. | Create clean Excel reports by making the unused part of a stacked bar transparent and removing its labels. | Customize label format to ‘0%’ and place it inside the bar for better readability.
// AI Prompts: Write C# code using Aspose.Cells to create a stacked bar chart, hide the second series, and show only percentage labels on the first series. | How can I format data labels as 0% and position them InsideEnd in an Aspose.Cells chart? | Explain steps to make a series transparent and suppress its data labels while enabling percentage labels on another series with Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartWithPercentageLabels
{
    // Creates an XLSX workbook, adds task data, builds a stacked bar chart used as a progress bar, makes the "Remaining" series transparent, disables its labels, and configures the "Completed" series to display only the percentage (0 % format) inside each bar.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a progress bar (stacked bar) chart
                // Column A – Category (Task)
                // Column B – Completed portion (value)
                // Column C – Remaining portion (value)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Completed");
                sheet.Cells["C1"].PutValue("Remaining");

                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(70);
                sheet.Cells["C2"].PutValue(30);

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["C3"].PutValue(55);

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(90);
                sheet.Cells["C4"].PutValue(10);

                // Add a stacked bar chart (used as a progress bar)
                // Note: In Aspose.Cells the stacked bar chart type is BarStacked
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the "Completed" series (visible part)
                int completedSeriesIdx = chart.NSeries.Add("B2:B4", true);
                // Add the "Remaining" series (invisible part)
                int remainingSeriesIdx = chart.NSeries.Add("C2:C4", true);

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the "Remaining" series so only the completed part is shown
                Series remainingSeries = chart.NSeries[remainingSeriesIdx];
                remainingSeries.Area.ForegroundColor = Color.Transparent; // make bar invisible

                // Disable all data label displays for the remaining series
                remainingSeries.DataLabels.ShowValue = false;
                remainingSeries.DataLabels.ShowPercentage = false;
                remainingSeries.DataLabels.ShowCategoryName = false;
                remainingSeries.DataLabels.ShowSeriesName = false;
                remainingSeries.DataLabels.ShowLegendKey = false;
                // Removed invalid ShowDataLabels property

                // Configure data labels for the visible (completed) series
                Series completedSeries = chart.NSeries[completedSeriesIdx];
                completedSeries.DataLabels.ShowPercentage = true;   // Show percentage value
                completedSeries.DataLabels.ShowValue = false;      // Hide raw value
                completedSeries.DataLabels.Position = LabelPositionType.InsideEnd; // Position inside the bar
                completedSeries.DataLabels.NumberFormat = "0%";    // Exact percentage format

                // Optional: set a distinct color for the completed portion
                completedSeries.Area.ForegroundColor = Color.Green;

                // Save the workbook
                string outputPath = "ProgressBarChartWithPercentageLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
