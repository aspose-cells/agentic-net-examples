// Title: Color Gantt Chart Bars by Priority with Aspose.Cells for C# (.NET)
// Description: Creates a workbook, adds task data with priority, builds a stacked‑bar Gantt chart, defines a custom palette, and colors each bar according to its priority before saving the file.
// Keywords: Aspose.Cells | C# | .NET | Gantt chart | custom chart colors | priority based coloring | stacked bar chart | ChangePalette | chart point formatting | conditional chart colors | project management visualization
// Common Searches: Aspose.Cells set bar colors by priority | C# Gantt chart custom palette Aspose.Cells | change workbook palette for chart colors .NET | conditional formatting of chart points Aspose.Cells | stacked bar Gantt chart with priority colors
// Developer Intent: Apply a priority‑driven color scheme to Gantt chart bars using a custom palette in Aspose.Cells for .NET.
// Use Cases: Show high, medium, and low priority tasks with distinct colors in a project timeline. | Maintain consistent priority colors across multiple charts in the same workbook. | Refresh bar colors automatically when the priority column is updated.
// AI Prompts: Write code that loads priority‑color mappings from a JSON file and applies them to Gantt chart points in Aspose.Cells. | Demonstrate how to replace solid bar colors with a gradient that reflects priority severity in a stacked bar chart. | Explain how to export the colored Gantt chart to PDF while preserving the custom palette colors.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartPriorityDemo
{
    // Creates a workbook, adds task data with priority, builds a stacked‑bar Gantt chart, defines a custom palette, and colors each bar according to its priority before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate task data: Task Name, Start Date, Duration (days), Priority
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("StartDate");
                sheet.Cells["C1"].PutValue("Duration");
                sheet.Cells["D1"].PutValue("Priority");

                // Sample tasks
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 7, 1));
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["D2"].PutValue("High");

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 7, 12));
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["D3"].PutValue("Medium");

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 8, 2));
                sheet.Cells["C4"].PutValue(8);
                sheet.Cells["D4"].PutValue("Low");

                // ------------------------------------------------------------
                // Add a Gantt‑like chart using a stacked bar chart
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 25, 10);
                Chart ganttChart = sheet.Charts[chartIndex];

                // First series: StartDate (in days)
                ganttChart.NSeries.Add("B2:B4", true);
                // Second series: Duration
                ganttChart.NSeries.Add("C2:C4", true);
                // Category (task names)
                ganttChart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Define custom colors for priorities
                // ------------------------------------------------------------
                Color highPriorityColor = Color.FromArgb(255, 102, 0);      // OrangeRed
                Color mediumPriorityColor = Color.FromArgb(255, 204, 0);    // Gold
                Color lowPriorityColor = Color.FromArgb(102, 204, 0);       // LightGreen

                // Ensure colors exist in the workbook palette (ChangePalette rule)
                workbook.ChangePalette(highPriorityColor, 0);
                workbook.ChangePalette(mediumPriorityColor, 1);
                workbook.ChangePalette(lowPriorityColor, 2);

                // ------------------------------------------------------------
                // Apply colors to each task bar based on priority
                // ------------------------------------------------------------
                // The second series (Duration) represents the visible bars.
                Series durationSeries = ganttChart.NSeries[1];

                for (int i = 0; i < durationSeries.Points.Count; i++)
                {
                    // Read priority from column D (index 3), rows start at 2
                    string priority = sheet.Cells[i + 2, 3].StringValue;

                    // Choose color based on priority
                    Color barColor = priority switch
                    {
                        "High" => highPriorityColor,
                        "Medium" => mediumPriorityColor,
                        "Low" => lowPriorityColor,
                        _ => Color.Gray
                    };

                    // Apply the color to the point's area
                    durationSeries.Points[i].Area.ForegroundColor = barColor;
                    durationSeries.Points[i].Area.Formatting = FormattingType.Custom;
                }

                // Optional: make colors varied (SeriesCollection.IsColorVaried rule)
                ganttChart.NSeries.IsColorVaried = true;

                // ------------------------------------------------------------
                // Save the workbook (lifecycle rule: save)
                // ------------------------------------------------------------
                workbook.Save("GanttChartPriorityDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
