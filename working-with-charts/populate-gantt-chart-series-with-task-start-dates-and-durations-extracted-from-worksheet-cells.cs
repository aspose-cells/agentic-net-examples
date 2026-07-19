// Title: Create a Gantt Chart in C# with Aspose.Cells Using Stacked Bar Series for Start Dates and Durations
// Description: This example builds a workbook, writes task names, start dates and duration values, defines cell ranges, adds a horizontal stacked‑bar chart, uses the start‑date series as a hidden offset, displays durations as visible bars, sets the task column as the category axis, tweaks the gap width, and saves the file as GanttChartDemo.xlsx.
// Keywords: Aspose.Cells | C# Gantt chart | stacked bar chart | start date offset | duration series | hide series Aspose.Cells | category axis data | Excel project timeline | automated chart generation | workbook chart binding
// Common Searches: Aspose.Cells Gantt chart example C# | how to hide a series in Aspose.Cells chart | populate chart series from worksheet cells Aspose.Cells | stacked bar chart with offset series C# | generate project schedule Excel using Aspose.Cells
// Developer Intent: Demonstrate how to bind worksheet ranges to a stacked bar chart, use the start‑date series as an invisible offset, and render task durations as Gantt bars.
// Use Cases: Automatically visualize a project schedule as a Gantt chart inside an Excel workbook. | Refresh the chart when new task rows are added without rewriting the code. | Export a ready‑to‑present timeline for status reports or stakeholder meetings.
// AI Prompts: Provide C# code that binds start‑date and duration ranges to a stacked bar chart in Aspose.Cells and hides the start series for offset purposes. | Show how to update the Gantt chart dynamically when additional tasks are appended, adjusting data ranges and series automatically. | Suggest styling options (colors, labels, date axis format) to improve the readability of an Aspose.Cells Gantt chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartDemo
{
    // This example builds a workbook, writes task names, start dates and duration values, defines cell ranges, adds a horizontal stacked‑bar chart, uses the start‑date series as a hidden offset, displays durations as visible bars, sets the task column as the category axis, tweaks the gap width, and saves the file as GanttChartDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                // Task 1
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 2));
                sheet.Cells["C2"].PutValue(5);

                // Task 2
                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 7));
                sheet.Cells["C3"].PutValue(10);

                // Task 3
                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 18));
                sheet.Cells["C4"].PutValue(4);

                // Determine the last row of data (adjust if data is dynamic)
                int lastRow = 4;

                // Define ranges for categories, start dates, and durations
                string taskRange = $"A2:A{lastRow}";
                string startRange = $"B2:B{lastRow}";
                string durationRange = $"C2:C{lastRow}";

                // Add a stacked bar chart (horizontal) – use BarStacked enum
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 26, 15);
                Chart chart = sheet.Charts[chartIndex];

                // First series: Start dates (used as offset). Hide its fill later.
                int startSeriesIdx = chart.NSeries.Add(startRange, true);
                Series startSeries = chart.NSeries[startSeriesIdx];
                startSeries.Name = "Start";

                // Second series: Durations (actual task lengths)
                int durationSeriesIdx = chart.NSeries.Add(durationRange, true);
                Series durationSeries = chart.NSeries[durationSeriesIdx];
                durationSeries.Name = "Duration";

                // Set the category (Y‑axis) to the task names
                chart.NSeries.CategoryData = taskRange;

                // Hide the start series from legend and display (acts as offset)
                startSeries.IsFiltered = true;

                // Adjust gap width for better visual spacing
                durationSeries.GapWidth = 150;

                // Save the workbook
                workbook.Save("GanttChartDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
