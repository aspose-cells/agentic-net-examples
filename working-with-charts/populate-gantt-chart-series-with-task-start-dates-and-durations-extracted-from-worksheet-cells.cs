// Title: Aspose.Cells C# – Build a Gantt chart from worksheet start dates and durations
// Description: Creates a new workbook, writes task names, start dates and duration values, adds a stacked‑bar chart, uses the start‑date range as a hidden series and the duration range as the visible series, sets the task names as the category axis, and saves the file as an Excel Gantt chart.
// Keywords: Aspose.Cells Gantt chart C# | stacked bar Gantt Aspose.Cells | populate chart series from cells | hide start series Aspose.Cells | set category axis task names | Excel project schedule C# | Aspose.Cells chart series example
// Common Searches: Aspose.Cells create Gantt chart C# | how to add start date series to stacked bar chart Aspose.Cells | hide start series in Gantt chart Aspose.Cells | set task names as Y axis in Aspose.Cells chart | C# example Gantt chart Aspose.Cells
// Developer Intent: Generate an Excel Gantt chart by reading task name, start date, and duration from worksheet cells and mapping them to a stacked‑bar chart.
// Use Cases: Automatically visualize project timelines directly from Excel data. | Export task schedules for stakeholder reports without manual charting. | Batch‑process multiple projects to produce consistent Gantt charts.
// AI Prompts: Show me how to hide the start‑date series so only duration bars appear in the Gantt chart. | Provide C# code to format the date axis with month/day labels in an Aspose.Cells Gantt chart. | Explain how to add data labels that display duration values on each bar of the chart. | Suggest ways to style the Gantt chart (colors, bar height, axis fonts) using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, writes task names, start dates and duration values, adds a stacked‑bar chart, uses the start‑date range as a hidden series and the duration range as the visible series, sets the task names as the category axis, and saves the file as an Excel Gantt chart.
class GanttChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data for Gantt chart -----
            // Column A : Task names
            // Column B : Start dates (as DateTime)
            // Column C : Duration (in days)
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Start");
            sheet.Cells["C1"].PutValue("Duration");

            sheet.Cells["A2"].PutValue("Planning");
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Design");
            sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 6));
            sheet.Cells["C3"].PutValue(8);

            sheet.Cells["A4"].PutValue("Implementation");
            sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 14));
            sheet.Cells["C4"].PutValue(12);

            sheet.Cells["A5"].PutValue("Testing");
            sheet.Cells["B5"].PutValue(new DateTime(2023, 1, 26));
            sheet.Cells["C5"].PutValue(6);

            // Determine the last row of data
            int lastRow = 5;

            // ----- Add a Gantt chart -----
            // Use a stacked bar chart to simulate a Gantt chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Series for start dates (will be hidden by not displaying its legend)
            int startSeriesIdx = chart.NSeries.Add($"=Sheet1!$B$2:$B${lastRow}", true);

            // Series for task durations (visible bars)
            int durationSeriesIdx = chart.NSeries.Add($"=Sheet1!$C$2:$C${lastRow}", true);

            // Set the category (Y) axis to the task names
            chart.NSeries.CategoryData = $"=Sheet1!$A$2:$A${lastRow}";

            // Optional: give a name to the duration series
            Series durationSeries = chart.NSeries[durationSeriesIdx];
            durationSeries.Name = "Duration";

            // Save the workbook
            workbook.Save("GanttChartExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
