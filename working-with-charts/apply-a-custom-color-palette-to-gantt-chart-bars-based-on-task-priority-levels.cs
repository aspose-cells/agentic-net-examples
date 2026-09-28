// Title: Create a Gantt chart with priority‑based custom colors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a stacked bar Gantt chart and assigns colors to each task bar according to a priority column (red for high, orange for medium, green for low). | Show how to make the start‑date series transparent in an Aspose.Cells stacked bar chart to achieve the classic Gantt view. | Provide a reusable method that reads a priority value from a worksheet cell and sets the corresponding chart point's foreground color.
// Common Searches: how to color Gantt chart bars by priority using Aspose.Cells C# | Aspose.Cells conditional bar colors based on worksheet column | C# create stacked bar chart for Gantt and hide start series Aspose | set individual point colors in Aspose.Cells chart series | custom color palette for task priority in Aspose.Cells Gantt chart
// Tags: apply custom palette to Aspose.Cells Gantt chart | hide start series in stacked bar chart Aspose | assign foreground color to chart points C# | priority driven bar colors Aspose.Cells | create Gantt chart with Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with task, start date, duration, and priority data, adds a stacked bar chart that mimics a Gantt chart, hides the start‑date series to make the bars appear as tasks, and then colors each duration bar red, orange, or green based on the priority value before saving the file as GanttChart_CustomPalette.xlsx.
class GanttChartWithCustomPalette
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: Task, Start Date, Duration (days), Priority (1=High,2=Medium,3=Low)
            string[,] data = new string[,]
            {
                { "Task", "Start", "Duration", "Priority" },
                { "Design", "2024-09-01", "5", "1" },
                { "Development", "2024-09-06", "10", "2" },
                { "Testing", "2024-09-16", "4", "3" },
                { "Deployment", "2024-09-20", "2", "1" }
            };

            // Populate worksheet with data
            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    sheet.Cells[r, c].PutValue(data[r, c]);
                }
            }

            // Convert Start column to DateTime and Duration to numeric
            int rows = data.GetLength(0);
            for (int r = 1; r < rows; r++)
            {
                DateTime start = DateTime.Parse(sheet.Cells[r, 1].StringValue);
                sheet.Cells[r, 1].PutValue(start);

                double duration = double.Parse(sheet.Cells[r, 2].StringValue);
                sheet.Cells[r, 2].PutValue(duration);
            }

            // Add a stacked bar chart (Gantt style)
            int chartRow = rows + 2;
            int chartCol = 0;
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, chartRow, chartCol, chartRow + 15, chartCol + 10);
            Chart ganttChart = sheet.Charts[chartIndex];

            // Define ranges for series and categories
            string startRange = $"B2:B{rows}";
            string durationRange = $"C2:C{rows}";
            string categoryRange = $"A2:A{rows}";

            // First series: Start (invisible)
            ganttChart.NSeries.Add(startRange, true);

            // Second series: Duration (colored based on priority)
            ganttChart.NSeries.Add(durationRange, true);

            // Set category axis labels (Task names)
            ganttChart.NSeries.CategoryData = categoryRange;

            // Hide the first series (Start) to create the Gantt effect
            ganttChart.NSeries[0].Area.ForegroundColor = Color.Transparent;
            ganttChart.NSeries[0].Area.Transparency = 100;

            // Define custom colors for priority levels
            Color highPriorityColor = Color.Red;
            Color mediumPriorityColor = Color.Orange;
            Color lowPriorityColor = Color.Green;

            // Apply colors to each bar based on priority
            for (int i = 0; i < ganttChart.NSeries[1].Points.Count; i++)
            {
                int priority = int.Parse(sheet.Cells[i + 1, 3].StringValue);
                Color barColor = lowPriorityColor; // default

                if (priority == 1)
                    barColor = highPriorityColor;
                else if (priority == 2)
                    barColor = mediumPriorityColor;

                ganttChart.NSeries[1].Points[i].Area.ForegroundColor = barColor;
            }

            // Save the workbook
            workbook.Save("GanttChart_CustomPalette.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
