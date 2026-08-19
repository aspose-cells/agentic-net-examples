// Title: Build an Excel Progress Bar Chart with Aspose.Cells for .NET (C#) Using a Stacked Bar and Hidden Series
// Description: A C# sample that creates a workbook, fills it with task names, a fixed total value, and actual progress numbers, then adds a horizontal stacked bar chart where the total series is concealed. The visible series is styled with custom fill and border colors and a zero‑gap width for a solid appearance, and the workbook is saved as ProgressBarChart.xlsx.
// Keywords: Aspose.Cells | C# Excel chart | progress bar visualization | stacked bar chart | hide series Aspose.Cells | gap width zero | custom series color | Excel automation .NET | chart formatting | export chart to PDF
// Common Searches: Aspose.Cells create progress bar chart C# | hide background series stacked bar Aspose.Cells | set gap width zero Aspose.Cells chart | customize series colors Aspose.Cells | export Aspose.Cells chart to PDF | C# tutorial Excel progress bar
// Developer Intent: Generate an Excel file that displays task completion as a progress bar by configuring a stacked bar chart with a concealed total series.
// Use Cases: Project dashboards showing percentage of work completed per task. | Quarterly reports visualizing sales target achievement as compact bars. | Operations summaries that illustrate resource utilization levels. | HR scorecards presenting employee performance metrics.
// AI Prompts: Write C# code with Aspose.Cells to create a progress‑bar style chart where colors and data ranges are configurable. | Explain how to hide a series in a stacked bar chart and set the gap width to zero for a solid bar effect using Aspose.Cells. | Show how to export the workbook containing the progress‑bar chart to PDF while preserving all chart formatting.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartDemo
{
    // A C# sample that creates a workbook, fills it with task names, a fixed total value, and actual progress numbers, then adds a horizontal stacked bar chart where the total series is concealed. The visible series is styled with custom fill and border colors and a zero‑gap width for a solid appearance, and the workbook is saved as ProgressBarChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Prepare data for the progress bar chart
                // Column A : Category (e.g., Task names)
                // Column B : Total value (fixed, e.g., 100)
                // Column C : Actual progress value
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Total");
                sheet.Cells["C1"].PutValue("Progress");

                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["A5"].PutValue("Task 4");

                // Total is constant 100 for all tasks
                for (int row = 2; row <= 5; row++)
                {
                    sheet.Cells[row - 1, 1].PutValue(100);
                }

                // Sample progress values
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(70);
                sheet.Cells["C4"].PutValue(55);
                sheet.Cells["C5"].PutValue(90);

                // -------------------------------------------------
                // Add a stacked bar chart (horizontal bars)
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series: Total (background) and Progress (foreground)
                // For a bar chart the series are added as row series (isColumnSeries = false)
                chart.NSeries.Add("B2:B5", false); // Series 0 – Total (will be hidden)
                chart.NSeries.Add("C2:C5", false); // Series 1 – Progress (visible)

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Hide the background series so only the progress part is shown
                chart.NSeries[0].IsFiltered = true; // Makes the series invisible

                // Remove gaps between bars for a solid look
                chart.NSeries[1].GapWidth = 0;

                // Optional: set the fill color of the progress series
                chart.NSeries[1].Area.ForegroundColor = Color.Green;

                // Optional: set the border of the progress series
                chart.NSeries[1].Border.Color = Color.DarkGreen;

                // Save the workbook to an Excel file
                string outputPath = "ProgressBarChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
