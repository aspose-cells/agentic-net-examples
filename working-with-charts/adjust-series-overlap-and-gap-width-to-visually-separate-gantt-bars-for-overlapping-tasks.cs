using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class GanttChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: Task name, Start (offset), Duration (length)
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Start");
        sheet.Cells["C1"].PutValue("Duration");

        string[] tasks = { "Task A", "Task B", "Task C" };
        int[] starts = { 1, 2, 3 };
        int[] durations = { 4, 3, 5 };

        for (int i = 0; i < tasks.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(tasks[i]);   // Column A
            sheet.Cells[i + 2, 1].PutValue(starts[i]); // Column B
            sheet.Cells[i + 2, 2].PutValue(durations[i]); // Column C
        }

        // Add a stacked bar chart (horizontal) to represent the Gantt chart
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // First series: Start offsets (make invisible so only Duration shows)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries[0].Values = "=B2:B4";
        chart.NSeries[0].Border.IsVisible = false;
        chart.NSeries[0].Area.ForegroundColor = Color.Transparent;
        chart.NSeries[0].Area.BackgroundColor = Color.Transparent;

        // Second series: Durations (visible bars)
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries[1].Values = "=C2:C4";

        // Set category (task) labels
        chart.NSeries.CategoryData = "A2:A4";

        // Adjust series overlap to separate overlapping bars (negative value)
        chart.NSeries[0].Overlap = -50;
        chart.NSeries[1].Overlap = -50;

        // Adjust gap width to increase spacing between task rows
        chart.NSeries[0].GapWidth = 200;
        chart.NSeries[1].GapWidth = 200;

        // Save the workbook with the configured chart
        workbook.Save("GanttChartOverlapGapWidth.xlsx");
    }
}