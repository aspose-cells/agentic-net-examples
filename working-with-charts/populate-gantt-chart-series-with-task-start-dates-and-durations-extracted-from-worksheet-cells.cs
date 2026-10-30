// Title: Generate a Gantt chart in C# with Aspose.Cells by binding task names, start dates and durations to a stacked‑bar chart series
// Common Searches: C# Aspose.Cells create Gantt chart from worksheet data | how to bind start date and duration columns to stacked bar series in Aspose.Cells | make invisible start series for offsetting bars in Aspose.Cells chart | set category axis labels from a range when building Gantt chart with Aspose.Cells | save Gantt chart as XLSX using Aspose.Cells .NET
// Tags: Aspose.Cells stacked bar Gantt chart C# | bind worksheet range to chart series Aspose.Cells | transparent start series Aspose.Cells | category axis labels from cell range Aspose.Cells | export Gantt chart to XLSX Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;

// The example creates a new workbook, writes task names, start dates and duration values, defines cell ranges, adds a stacked‑bar chart, adds an invisible start‑date series to offset the visible duration series, assigns the task names as Y‑axis category labels, and saves the file as GanttChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Tasks";

            // Header
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Start");
            sheet.Cells["C1"].PutValue("Duration");

            // Sample data
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
            sheet.Cells["C3"].PutValue(8);

            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 5));
            sheet.Cells["C4"].PutValue(3);

            // Determine the last row with data
            int lastRow = sheet.Cells.MaxDataRow;
            if (lastRow < 1) return; // No data to plot

            // Define ranges for categories, start dates, and durations
            string categoryRange = $"A2:A{lastRow}";
            string startRange = $"B2:B{lastRow}";
            string durationRange = $"C2:C{lastRow}";

            // Add a stacked bar chart (used for Gantt charts)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 25, 10);
            Chart ganttChart = sheet.Charts[chartIndex];
            ganttChart.Title.Text = "Gantt Chart";

            // Series 1: Start dates (invisible, used to offset the duration bars)
            int startSeriesIdx = ganttChart.NSeries.Add(startRange, true);
            Series startSeries = ganttChart.NSeries[startSeriesIdx];
            startSeries.Name = "Start";
            startSeries.Area.ForegroundColor = Color.Transparent;
            startSeries.Border.IsVisible = false;

            // Series 2: Durations (visible bars)
            int durationSeriesIdx = ganttChart.NSeries.Add(durationRange, true);
            Series durationSeries = ganttChart.NSeries[durationSeriesIdx];
            durationSeries.Name = "Duration";
            durationSeries.Area.ForegroundColor = Color.Blue;

            // Assign task names as category (Y‑axis) labels
            ganttChart.NSeries.CategoryData = categoryRange;

            // Optional: reverse the category axis (not supported in this version)
            // ganttChart.CategoryAxis.IsReverse = true;

            // Save the workbook
            workbook.Save("GanttChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
