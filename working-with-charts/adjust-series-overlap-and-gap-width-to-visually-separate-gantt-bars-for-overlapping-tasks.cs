// Title: Set Series Overlap & Gap Width to Separate Overlapping Gantt Bars – Aspose.Cells C# Example
// Description: This C# sample creates an Excel workbook, populates it with task, start, and duration data, adds a stacked bar chart for a Gantt view, hides the start series, and configures the Duration series with a negative Overlap and a custom GapWidth to clearly separate overlapping bars before saving the file.
// Keywords: Aspose.Cells | C# | Gantt chart | stacked bar chart | series overlap | negative overlap | gap width | chart spacing | transparent series | Excel chart formatting | overlapping tasks visualization
// Common Searches: Aspose.Cells set negative series overlap | change gap width stacked bar chart Aspose.Cells | separate overlapping Gantt bars C# | hide start series in Aspose.Cells chart | adjust bar spacing in Excel using Aspose.Cells
// Developer Intent: Apply the Overlap and GapWidth properties of a chart series to visually separate bars that would otherwise overlap in a Gantt chart.
// Use Cases: Display tasks with overlapping time periods as distinct bars in a Gantt chart. | Hide positioning data (Start) while keeping it for bar alignment by making the series transparent. | Fine‑tune vertical spacing between task rows to improve readability of stacked bar charts.
// AI Prompts: Write C# code using Aspose.Cells that creates a Gantt chart and sets series Overlap to -40 and GapWidth to 150. | Explain how Overlap and GapWidth affect stacked bar charts in Aspose.Cells, with recommended value ranges. | Provide a step‑by‑step guide to make a chart series invisible but still used for positioning in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# sample creates an Excel workbook, populates it with task, start, and duration data, adds a stacked bar chart for a Gantt view, hides the start series, and configures the Duration series with a negative Overlap and a custom GapWidth to clearly separate overlapping bars before saving the file.
class GanttChartOverlapDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare sample data for a Gantt chart (Task, Start, Duration)
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Start");
        sheet.Cells["C1"].PutValue("Duration");

        string[] tasks = { "Task A", "Task B", "Task C" };
        int[] starts = { 1, 2, 3 };
        int[] durations = { 5, 4, 6 };

        for (int i = 0; i < tasks.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(tasks[i]);   // Column A
            sheet.Cells[i + 2, 1].PutValue(starts[i]); // Column B
            sheet.Cells[i + 2, 2].PutValue(durations[i]); // Column C
        }

        // Add a stacked bar chart (commonly used for Gantt charts)
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // First series: Start (invisible, used only for positioning)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries[0].Values = "=B2:B4";
        chart.NSeries[0].Name = "Start";
        // Make the start series transparent so it doesn't appear in the legend or chart
        chart.NSeries[0].Area.ForegroundColor = Color.Transparent;
        chart.NSeries[0].IsFiltered = true; // hide from legend

        // Second series: Duration (visible bars)
        chart.NSeries.Add("C2:C4", false);
        chart.NSeries[1].Values = "=C2:C4";
        chart.NSeries[1].Name = "Duration";

        // Set category (Y‑axis) labels to the task names
        chart.NSeries.CategoryData = "A2:A4";

        // Adjust series overlap and gap width to separate overlapping bars
        chart.NSeries[1].Overlap = -30;   // Negative value separates bars (-100 to 100)
        chart.NSeries[1].GapWidth = 100; // Reduces space between clusters (0 to 500)

        // Save the workbook with the chart
        workbook.Save("GanttOverlapDemo.xlsx");
    }
}
