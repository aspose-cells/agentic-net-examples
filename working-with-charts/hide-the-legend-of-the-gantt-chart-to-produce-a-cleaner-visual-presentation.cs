// Title: Hide Chart Legend in a Gantt‑Style Stacked Bar Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts task names with start/end dates, builds a stacked bar chart that mimics a Gantt chart, disables its legend via Chart.ShowLegend = false, and saves the workbook.
// Keywords: Aspose.Cells | C# | hide chart legend | remove legend Excel chart | Gantt chart Aspose.Cells | stacked bar chart | Chart.ShowLegend false | project timeline Excel | Excel automation .NET | chart formatting Aspose
// Common Searches: Aspose.Cells hide legend C# | How to remove legend from Gantt chart using Aspose.Cells | Chart.ShowLegend false example | C# code to hide Excel chart legend Aspose | Hide legend in stacked bar chart Aspose.Cells
// Developer Intent: The developer wants to remove the legend from a Gantt‑style stacked bar chart to achieve a cleaner visual presentation.
// Use Cases: Generate a project timeline workbook where the legend adds no value, improving stakeholder readability. | Create multiple Gantt charts in a single report and hide all legends to keep the layout compact. | Export the chart to PDF or image without a legend, reducing visual clutter in documentation.
// AI Prompts: Write C# code that uses Aspose.Cells to hide the legend of an existing chart in a workbook. | Explain how to toggle the ShowLegend property conditionally based on user preferences in Aspose.Cells. | Provide a method that hides legends for all charts in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttLegendHide
{
    // Creates a workbook, inserts task names with start/end dates, builds a stacked bar chart that mimics a Gantt chart, disables its legend via Chart.ShowLegend = false, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a Gantt‑like chart
                // Task names
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["A4"].PutValue("Testing");

                // Start dates
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["B3"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["B4"].PutValue(new DateTime(2023, 3, 15));

                // End dates
                sheet.Cells["C1"].PutValue("End");
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 31));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 3, 14));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 4, 30));

                // Add a stacked bar chart (used as a Gantt‑style chart)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range: start and end dates
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the legend for a cleaner visual presentation
                chart.ShowLegend = false;

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
