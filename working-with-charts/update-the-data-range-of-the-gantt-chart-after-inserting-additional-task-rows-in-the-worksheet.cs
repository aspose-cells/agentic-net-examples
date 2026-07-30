// Title: Aspose.Cells for .NET – Expand Gantt‑style chart after adding task rows
// Description: Shows how to build a workbook with a task table, create a stacked‑bar chart that emulates a Gantt chart, insert additional task rows, resize the ListObject, update the chart’s data and category ranges, recalculate, and save the workbook.
// Keywords: Aspose.Cells | C# | Gantt chart | stacked bar chart | update chart range | resize ListObject | insert rows Excel | dynamic Excel chart | programmatic Excel automation | Excel workbook generation
// Common Searches: how to extend a Gantt chart after adding rows with Aspose.Cells | C# resize ListObject and refresh chart data range | Aspose.Cells stacked bar chart category range update | programmatically grow Excel chart source range | add tasks to Gantt‑style chart using Aspose.Cells for .NET
// Developer Intent: Refresh a Gantt‑style chart to include newly inserted task rows by resizing the source table and adjusting the chart’s data and category ranges.
// Use Cases: Append project tasks to an existing Gantt chart without recreating the chart. | Automatically expand a ListObject and keep a linked stacked‑bar chart in sync. | Generate a schedule workbook where tasks can be added on‑the‑fly and the visualization updates.
// AI Prompts: Write C# code with Aspose.Cells that inserts new task rows, resizes a ListObject, and updates a stacked‑bar Gantt‑style chart’s data range. | Show how to change the category range of a Gantt‑like chart after adding rows to the source table in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace GanttChartUpdateDemo
{
    // Shows how to build a workbook with a task table, create a stacked‑bar chart that emulates a Gantt chart, insert additional task rows, resize the ListObject, update the chart’s data and category ranges, recalculate, and save the workbook.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- 1. Populate initial task data ----------
                // Headers
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Finish");

                // Sample tasks (rows 2‑5)
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 5));

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 1, 8));

                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 1, 6));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 1, 10));

                sheet.Cells["A5"].PutValue("Task 4");
                sheet.Cells["B5"].PutValue(new DateTime(2023, 1, 9));
                sheet.Cells["C5"].PutValue(new DateTime(2023, 1, 12));

                // ---------- 2. Create a table (ListObject) for the data ----------
                // Table covers A1:C5 (including headers)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject taskTable = sheet.ListObjects[tableIndex];
                taskTable.DisplayName = "TaskTable";

                // ---------- 3. Add a Gantt‑like chart linked to the table ----------
                // Use a stacked bar chart to emulate a Gantt chart (ChartType.Gantt is not available in this version)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 3, 19, 11);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Initial data range (A1:C5). For Gantt‑like charts the series data is usually the start and finish columns.
                // The first column (Task) is used for categories.
                ganttChart.SetChartDataRange("A1:C5", true);
                ganttChart.NSeries.CategoryData = "A2:A5";

                // ---------- 4. Insert additional task rows ----------
                // Insert two new rows after the existing data (after row 5)
                sheet.Cells.InsertRows(5, 2); // rows are zero‑based, so row index 5 is the 6th row

                // Fill the newly inserted rows with task data
                sheet.Cells["A6"].PutValue("Task 5");
                sheet.Cells["B6"].PutValue(new DateTime(2023, 1, 11));
                sheet.Cells["C6"].PutValue(new DateTime(2023, 1, 15));

                sheet.Cells["A7"].PutValue("Task 6");
                sheet.Cells["B7"].PutValue(new DateTime(2023, 1, 13));
                sheet.Cells["C7"].PutValue(new DateTime(2023, 1, 18));

                // ---------- 5. Resize the table to include the new rows ----------
                // New end row index is 6 (zero‑based) because we now have rows 0‑6 with data (7 rows total)
                taskTable.Resize(0, 0, 6, 2, true);

                // ---------- 6. Update the chart data range ----------
                // New range now spans A1:C7
                ganttChart.SetChartDataRange("A1:C7", true);
                ganttChart.NSeries.CategoryData = "A2:A7";

                // Re‑calculate the chart to apply the changes
                ganttChart.Calculate();

                // ---------- 7. Save the workbook ----------
                string outputPath = "UpdatedGanttChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
