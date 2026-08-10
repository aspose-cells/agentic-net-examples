// Title: Aspose.Cells .NET – Adjust Series Overlap & Gap Width to Separate Gantt Bars
// Description: C# example that builds a horizontal stacked‑bar Gantt chart, hides the start‑date series, and applies Overlap = -40 and GapWidth = 150 to both series (and the chart) so overlapping tasks are visually separated. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells series overlap | gap width stacked bar chart | Gantt chart spacing .NET | horizontal stacked bar Aspose | Excel chart formatting C# | visual separation Gantt bars | Aspose.Cells chart properties
// Common Searches: set series overlap Aspose.Cells | gap width for stacked bar chart .NET | separate overlapping Gantt tasks Excel | Aspose.Cells horizontal Gantt example | adjust chart spacing Aspose.Cells C#
// Developer Intent: Modify a stacked‑bar Gantt chart in Aspose.Cells so that overlapping task bars are spaced apart by configuring the Overlap and GapWidth properties of the series and chart.
// Use Cases: Create a horizontal stacked‑bar Gantt chart where the start‑date series is invisible and the duration series forms the visible bars. | Apply the same Overlap (-40) and GapWidth (150) settings to multiple series to ensure consistent spacing across all tasks. | Generate an Excel workbook with clearly spaced Gantt bars for project‑timeline reporting.
// AI Prompts: Write C# code using Aspose.Cells to build a Gantt chart and set series Overlap and GapWidth for visual separation. | Explain the impact of Overlap and GapWidth on stacked‑bar charts in Aspose.Cells with code snippets. | Provide debugging steps when Gantt bars still overlap after setting Overlap and GapWidth in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace GanttChartExample
{
    // C# example that builds a horizontal stacked‑bar Gantt chart, hides the start‑date series, and applies Overlap = -40 and GapWidth = 150 to both series (and the chart) so overlapping tasks are visually separated. The workbook is saved as an Excel file.
    class Program
    {
        static void Main()
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

                // Sample data – store start dates as Excel serial numbers (OADate)
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(DateTime.Today.ToOADate());
                sheet.Cells["C2"].PutValue(5);

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(DateTime.Today.AddDays(3).ToOADate());
                sheet.Cells["C3"].PutValue(7);

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(DateTime.Today.AddDays(6).ToOADate());
                sheet.Cells["C4"].PutValue(4);

                // Add a stacked bar chart (horizontal) to represent the Gantt bars
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // First series: Start dates (invisible, used for positioning)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Values = "=B2:B4";
                chart.NSeries[0].Name = "Start";
                chart.NSeries[0].Border.IsVisible = false;
                chart.NSeries[0].Area.ForegroundColor = Color.Transparent;

                // Second series: Duration (visible Gantt bars)
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Values = "=C2:C4";
                chart.NSeries[1].Name = "Duration";

                // Set category (task) labels
                chart.NSeries.CategoryData = "A2:A4";

                // Adjust overlap and gap width for better visual separation
                chart.NSeries[0].Overlap = -40;
                chart.NSeries[1].Overlap = -40;
                chart.NSeries[0].GapWidth = 150;
                chart.NSeries[1].GapWidth = 150;
                chart.GapWidth = 150;

                // Save the workbook
                workbook.Save("GanttChartWithOverlapAndGapWidth.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
