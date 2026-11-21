// Title: Set localized chart title and axis labels on a stacked‑bar Gantt chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart title with Chinese text | How to add axis titles to a stacked bar chart in Aspose.Cells .NET | Localizing chart labels for a Gantt‑style chart with Aspose.Cells | Set bold font for chart title and axes using Aspose.Cells C# | Create a Gantt chart using stacked bar chart and set localized axis titles Aspose.Cells
// Tags: localized chart title Aspose.Cells | stacked bar chart axis titles .NET | simulate Gantt chart with Aspose.Cells | chart title bold font C# | Chinese chart labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a stacked bar chart to mimic a Gantt chart, assigns Chinese strings to the chart title, category axis, and value axis, applies bold styling to each title, and saves the file as GanttChart.xlsx.
class GanttChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Aspose.Cells does not have a dedicated Gantt chart type in the current version.
            // Use a stacked bar chart as a basis for creating a Gantt-like representation.
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart ganttChart = sheet.Charts[chartIndex];

            // Localized strings (replace with actual localized resources as needed)
            string chartTitle = "项目进度图";          // Chinese title
            string categoryAxisTitle = "任务名称";    // Chinese axis label
            string valueAxisTitle = "时间线";        // Chinese axis label

            // Set the chart title
            ganttChart.Title.Text = chartTitle;
            ganttChart.Title.Font.IsBold = true;

            // Set the category (vertical) axis title
            ganttChart.CategoryAxis.Title.Text = categoryAxisTitle;
            ganttChart.CategoryAxis.Title.Font.IsBold = true;

            // Set the value (horizontal) axis title
            ganttChart.ValueAxis.Title.Text = valueAxisTitle;
            ganttChart.ValueAxis.Title.Font.IsBold = true;

            // (Optional) Populate data for the Gantt chart here
            // Example: sheet.Cells["A1"].PutValue("Task");
            // ...

            // Save the workbook to a file
            workbook.Save("GanttChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
