// Title: Color‑code Gantt chart bars by task priority using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set bar color in Gantt chart based on priority column | How to apply conditional colors to stacked bar chart points with Aspose.Cells | Change Gantt chart task bar colors programmatically in .NET | C# example for coloring Gantt chart bars according to priority values | Aspose.Cells assign different colors to series points in a stacked bar chart
// Tags: set series point color Aspose.Cells C# | conditional bar coloring Gantt chart Aspose.Cells | stacked bar chart color customization .NET | priority based Gantt chart colors Excel | apply custom palette to chart series Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;

// The sample creates an Excel workbook, adds task data with start dates, durations, and priority levels, builds a stacked bar chart to act as a Gantt chart, hides the start‑date series, then iterates over the duration series points and assigns red, orange, or green colors based on the priority value before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "GanttData";

            // ---------- Populate sample data ----------
            // Headers
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Start");
            sheet.Cells["C1"].PutValue("Duration");
            sheet.Cells["D1"].PutValue("Priority");

            // Row 2
            sheet.Cells["A2"].PutValue("Design");
            sheet.Cells["B2"].PutValue(new DateTime(2023, 10, 1));
            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["D2"].PutValue("High");

            // Row 3
            sheet.Cells["A3"].PutValue("Development");
            sheet.Cells["B3"].PutValue(new DateTime(2023, 10, 6));
            sheet.Cells["C3"].PutValue(10);
            sheet.Cells["D3"].PutValue("Medium");

            // Row 4
            sheet.Cells["A4"].PutValue("Testing");
            sheet.Cells["B4"].PutValue(new DateTime(2023, 10, 16));
            sheet.Cells["C4"].PutValue(4);
            sheet.Cells["D4"].PutValue("Low");

            // Format the Start column as dates
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Excel date format
            sheet.Cells["B2:B4"].SetStyle(dateStyle);

            // ---------- Add a stacked bar chart to act as a Gantt chart ----------
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 25, 10);
            Chart ganttChart = sheet.Charts[chartIndex];
            ganttChart.Title.Text = "Project Gantt";

            // Series 1: Start dates (invisible, used for offset)
            ganttChart.NSeries.Add("B2:B4", true);
            ganttChart.NSeries[0].Area.ForegroundColor = Color.Transparent; // hide
            ganttChart.NSeries[0].Border.IsVisible = false;

            // Series 2: Duration (visible bars)
            ganttChart.NSeries.Add("C2:C4", true);
            ganttChart.NSeries[1].PlotOnSecondAxis = false;

            // Set categories (task names)
            ganttChart.NSeries.CategoryData = "A2:A4";

            // ---------- Apply custom colors based on priority ----------
            // Priority column is D (zero‑based index 3)
            for (int i = 0; i < ganttChart.NSeries[1].Points.Count; i++)
            {
                string priority = sheet.Cells[i + 1, 3].StringValue; // row offset + column D
                Color barColor = priority switch
                {
                    "High"   => Color.Red,
                    "Medium" => Color.Orange,
                    "Low"    => Color.Green,
                    _        => Color.Gray
                };
                ganttChart.NSeries[1].Points[i].Area.ForegroundColor = barColor;
            }

            // ---------- Save the workbook ----------
            workbook.Save("GanttChart.xlsx");                       // save
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
