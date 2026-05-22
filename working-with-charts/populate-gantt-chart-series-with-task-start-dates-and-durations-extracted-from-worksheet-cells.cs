using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: Task name, Start Date, Duration (in days)
            // Header
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("StartDate");
            sheet.Cells["C1"].PutValue("Duration");

            // Data rows
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["C2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
            sheet.Cells["C3"].PutValue(3);

            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 5));
            sheet.Cells["C4"].PutValue(7);

            sheet.Cells["A5"].PutValue("Task 4");
            sheet.Cells["B5"].PutValue(new DateTime(2023, 1, 2));
            sheet.Cells["C5"].PutValue(4);

            // Add a stacked bar chart (used for Gantt representation)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 26, 15);
            Chart chart = sheet.Charts[chartIndex];

            // First series: Start dates (will be hidden)
            int startSeriesIdx = chart.NSeries.Add("B2:B5", true);
            Series startSeries = chart.NSeries[startSeriesIdx];
            startSeries.Name = "Start";
            // Hide the start series so only the duration appears
            startSeries.IsFiltered = true;

            // Second series: Durations
            int durationSeriesIdx = chart.NSeries.Add("C2:C5", true);
            Series durationSeries = chart.NSeries[durationSeriesIdx];
            durationSeries.Name = "Duration";

            // Set the category (Y‑axis) to the task names
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: set chart title
            chart.Title.Text = "Project Gantt Chart";

            // Save the workbook
            workbook.Save("GanttChartOutput.xlsx");
        }
    }
}