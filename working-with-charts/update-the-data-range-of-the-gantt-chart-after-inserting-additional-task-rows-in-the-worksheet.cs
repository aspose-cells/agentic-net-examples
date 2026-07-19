// Title: Aspose.Cells .NET – Expand Gantt Chart Range After Inserting Task Rows
// Description: Demonstrates how to add new task rows to a worksheet, compute the last populated row with Worksheet.Cells.MaxDataRow, rebuild the address string, call SetChartDataRange to include the new rows, recalculate the stacked‑bar (Gantt) chart, and save the workbook.
// Keywords: Aspose.Cells C# Gantt chart | SetChartDataRange dynamic range | Worksheet.Cells.MaxDataRow | add rows to stacked bar chart | .NET chart data refresh | recalculate chart Aspose.Cells | Excel Gantt chart update
// Common Searches: how to extend Gantt chart range after adding rows Aspose.Cells | C# update chart data range when inserting rows | Aspose.Cells refresh stacked bar chart after data change | get last data row for chart range Aspose.Cells | dynamic chart range in .NET Excel library
// Developer Intent: Automatically include newly inserted task rows in an existing Gantt chart by recalculating its data range.
// Use Cases: Insert additional tasks and let the chart grow without hard‑coding the range. | Use MaxDataRow to create a range that adapts to any number of tasks. | Recalculate the chart after changing the range to ensure the saved file displays the updated Gantt view.
// AI Prompts: Write C# code that adds several task rows to a worksheet and updates the Gantt chart range using Aspose.Cells. | Show how to use Worksheet.Cells.MaxDataRow to build a dynamic chart address for a stacked‑bar Gantt chart. | Explain the steps to recalculate and save a workbook after modifying a chart's data range with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartUpdateDemo
{
    // Demonstrates how to add new task rows to a worksheet, compute the last populated row with Worksheet.Cells.MaxDataRow, rebuild the address string, call SetChartDataRange to include the new rows, recalculate the stacked‑bar (Gantt) chart, and save the workbook.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Populate initial task data -----
                // Header row
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("StartDate");

                // Sample tasks (rows 2‑4)
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 5));
                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 10));

                // ----- Add a Gantt chart -----
                // In Aspose.Cells a Gantt chart is created as a stacked bar chart
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 8);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Initial data range (tasks and start dates)
                // The range includes the header row; the boolean 'true' indicates vertical series (by column)
                ganttChart.SetChartDataRange("A1:B4", true);

                // ----- Insert additional task rows -----
                // Insert two new rows after the existing tasks (after row 4)
                sheet.Cells.InsertRows(4, 2); // rows are zero‑based; this inserts rows 5‑6

                // Populate the newly inserted rows with new tasks
                sheet.Cells["A5"].PutValue("Task 4");
                sheet.Cells["B5"].PutValue(new DateTime(2023, 1, 15));
                sheet.Cells["A6"].PutValue("Task 5");
                sheet.Cells["B6"].PutValue(new DateTime(2023, 1, 20));

                // ----- Update the chart data range to include the new rows -----
                // Determine the last row that contains data
                int lastDataRow = sheet.Cells.MaxDataRow; // zero‑based index
                // Build the new address string (add 1 because Excel addresses are 1‑based)
                string newRange = $"A1:B{lastDataRow + 1}";
                ganttChart.SetChartDataRange(newRange, true);

                // Recalculate the chart to apply the updated range
                ganttChart.Calculate();

                // Save the workbook
                workbook.Save("GanttChartUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
