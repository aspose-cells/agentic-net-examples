// Title: Aspose.Cells C# – Update Gantt Chart Data Range After Adding Task Rows
// Description: C# example that creates a workbook, builds a stacked‑bar Gantt chart, inserts additional task rows, and expands the chart’s data range using SetChartDataRange followed by Chart.Calculate to refresh the visual.
// Keywords: Aspose.Cells | C# | .NET | Gantt chart | stacked bar chart | SetChartDataRange | InsertRows | update chart range | refresh chart | Excel automation | dynamic tasks
// Common Searches: Aspose.Cells update chart range after inserting rows | C# expand Gantt chart data range | SetChartDataRange example Aspose.Cells | Refresh stacked bar chart in .NET | Add tasks to Gantt chart programmatically | Aspose.Cells dynamic chart data source
// Developer Intent: Adjust an existing Gantt‑style chart to include newly inserted task rows.
// Use Cases: Insert new task rows, fill start dates and durations, then call SetChartDataRange with the extended range. | Recalculate the chart after changing the data range to display added tasks instantly. | Generate an up‑to‑date Gantt chart workbook after dynamic modifications to the task list.
// AI Prompts: Write C# code that adds multiple task rows to an Aspose.Cells worksheet and automatically updates a stacked‑bar Gantt chart. | Show how to use Chart.Calculate after SetChartDataRange to refresh a Gantt chart in Aspose.Cells. | Explain the steps to programmatically adjust a Gantt chart’s series range when tasks are added at runtime.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttUpdateDemo
{
    // C# example that creates a workbook, builds a stacked‑bar Gantt chart, inserts additional task rows, and expands the chart’s data range using SetChartDataRange followed by Chart.Calculate to refresh the visual.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate initial task data for the Gantt chart
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                // Sample tasks (3 rows)
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(5); // days

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
                sheet.Cells["C3"].PutValue(8);

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 5));
                sheet.Cells["C4"].PutValue(4);

                // -------------------------------------------------
                // 2. Add a Gantt‑like chart based on the initial data
                // -------------------------------------------------
                // Use a stacked bar chart to emulate a Gantt chart
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 7);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Set the data range (including header row)
                // The chart expects series data in columns B (Start) and C (Duration)
                ganttChart.SetChartDataRange("A1:C4", true);

                // -------------------------------------------------
                // 3. Insert additional task rows below the existing ones
                // -------------------------------------------------
                // Insert 2 new rows after row 4 (zero‑based index 4)
                sheet.Cells.InsertRows(4, 2, true);

                // Populate the newly inserted rows with new tasks
                sheet.Cells["A5"].PutValue("Task 4");
                sheet.Cells["B5"].PutValue(new DateTime(2023, 1, 7));
                sheet.Cells["C5"].PutValue(6);

                sheet.Cells["A6"].PutValue("Task 5");
                sheet.Cells["B6"].PutValue(new DateTime(2023, 1, 9));
                sheet.Cells["C6"].PutValue(3);

                // -------------------------------------------------
                // 4. Update the chart data range to include the new rows
                // -------------------------------------------------
                ganttChart.SetChartDataRange("A1:C6", true);
                ganttChart.Calculate(); // Refresh layout

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("GanttChartUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
