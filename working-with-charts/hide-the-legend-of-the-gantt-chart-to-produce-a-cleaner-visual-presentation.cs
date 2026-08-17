// Title: Hide Chart Legend in a Gantt‑Style Stacked Bar Chart with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, add task data, create a stacked‑bar Gantt chart, and suppress the chart legend by setting Chart.ShowLegend = false before saving as GanttChart_NoLegend.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Gantt chart | stacked bar chart | hide legend | Chart.ShowLegend | Excel chart formatting | remove chart legend | Aspose.Cells API
// Common Searches: Aspose.Cells hide legend C# | remove legend from Gantt chart Aspose | Chart.ShowLegend false example | C# generate Gantt chart without legend | Aspose.Cells chart formatting tutorial
// Developer Intent: The developer wants to remove the legend from a Gantt‑style stacked bar chart generated with Aspose.Cells in a C# .NET project.
// Use Cases: Create project‑timeline charts that match corporate slide templates without a legend. | Automate batch generation of Gantt charts where legends add unnecessary visual clutter. | Export Excel reports for presentations that display only task bars.
// AI Prompts: Write C# code using Aspose.Cells to hide the legend of a specific chart while keeping other settings unchanged. | Provide a reusable method that toggles Chart.ShowLegend based on a boolean argument when building a Gantt chart. | Show how to loop through all worksheets and hide legends for every chart in a workbook with Aspose.Cells. | Explain how hiding chart legends affects Excel file size and rendering performance.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttLegendHide
{
    // Shows how to build a workbook, add task data, create a stacked‑bar Gantt chart, and suppress the chart legend by setting Chart.ShowLegend = false before saving as GanttChart_NoLegend.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a Gantt-like chart
                // Columns: Task, Start Date, End Date
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("End");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 10));

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 11));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 15));

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 16));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 3, 5));

                // Add a stacked bar chart (used to emulate a Gantt chart)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range: start dates and end dates
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the legend for a cleaner visual presentation
                chart.ShowLegend = false;

                // Save the workbook
                string outputPath = "GanttChart_NoLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
