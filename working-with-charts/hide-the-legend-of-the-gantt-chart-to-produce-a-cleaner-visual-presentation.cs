// Title: Hide the Legend in a Gantt‑Style Stacked Bar Chart Using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds task data, builds a stacked bar chart that mimics a Gantt chart, makes the start‑date series transparent, assigns task names as categories, disables the chart legend via the ShowLegend property, and saves the file as GanttChart_NoLegend.xlsx.
// Keywords: Aspose.Cells hide chart legend | C# hide legend Aspose.Cells | Aspose.Cells Gantt chart legend | .NET stacked bar chart legend | disable chart legend Excel | Aspose.Cells chart formatting | remove legend Aspose.Cells US | Aspose.Cells global chart customization
// Common Searches: how to hide legend in Aspose.Cells chart | Aspose.Cells .NET remove Gantt chart legend | disable legend for stacked bar chart C# | Aspose.Cells ShowLegend false example | hide Excel chart legend programmatically
// Developer Intent: The developer needs to suppress the legend of a Gantt‑style stacked bar chart to achieve a cleaner visual output.
// Use Cases: Generate project‑timeline Gantt charts without unnecessary legends. | Create compact Excel reports where chart legends waste space. | Design dashboard visuals that rely on color cues instead of legends.
// AI Prompts: Write C# code with Aspose.Cells that hides the legend of any chart type. | Explain how to toggle the ShowLegend property based on user settings in Aspose.Cells. | Show how to hide a chart legend and then customize axis titles and data labels in a Gantt chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttLegendHide
{
    // This example creates a workbook, adds task data, builds a stacked bar chart that mimics a Gantt chart, makes the start‑date series transparent, assigns task names as categories, disables the chart legend via the ShowLegend property, and saves the file as GanttChart_NoLegend.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a Gantt‑like chart
                // Columns: Task, Start Date, End Date
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("End");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 15));

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 16));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 28));

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 3, 15));

                // Add a stacked bar chart (used to emulate a Gantt chart)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Series 1 – Start dates (invisible, used for positioning)
                ganttChart.NSeries.Add("B2:B4", true);
                ganttChart.NSeries[0].IsColorVaried = false;
                ganttChart.NSeries[0].PlotOnSecondAxis = false;
                // Hide the first series by making it transparent
                ganttChart.NSeries[0].Area.ForegroundColor = System.Drawing.Color.Transparent;

                // Series 2 – End dates (duration)
                ganttChart.NSeries.Add("C2:C4", true);
                ganttChart.NSeries[1].IsColorVaried = true;

                // Category (task names)
                ganttChart.NSeries.CategoryData = "A2:A4";

                // Hide the legend for a cleaner visual presentation
                ganttChart.ShowLegend = false;

                // Save the workbook
                workbook.Save("GanttChart_NoLegend.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
