using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace GanttChartUpdateDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate initial Gantt data (Task, Start, End)
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("End");

                // Sample tasks
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(DateTime.Today);
                sheet.Cells["C2"].PutValue(DateTime.Today.AddDays(3));

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(DateTime.Today.AddDays(1));
                sheet.Cells["C3"].PutValue(DateTime.Today.AddDays(4));

                // -------------------------------------------------
                // 2. Create a table (ListObject) for the data range
                // -------------------------------------------------
                int tableIdx = sheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject table = sheet.ListObjects[tableIdx];
                table.DisplayName = "GanttData";

                // -------------------------------------------------
                // 3. Add a Gantt chart based on the table data
                // -------------------------------------------------
                // Gantt chart is a stacked bar chart
                int chartIdx = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 8);
                Chart ganttChart = sheet.Charts[chartIdx];

                // Helper column for Duration (End - Start)
                sheet.Cells["D1"].PutValue("Duration");
                for (int r = 2; r <= 3; r++)
                {
                    sheet.Cells[r - 1, 3].Formula = $"=C{r}-B{r}";
                }

                // Add series: Start (invisible) and Duration (visible)
                ganttChart.NSeries.Add("B2:B3", true); // Start
                ganttChart.NSeries[0].IsColorVaried = false; // keep start series invisible later
                ganttChart.NSeries.Add("D2:D3", true); // Duration
                ganttChart.NSeries[1].IsColorVaried = true;

                // Set category (Task names)
                ganttChart.NSeries.CategoryData = "A2:A3";

                // -------------------------------------------------
                // 4. Insert additional task rows
                // -------------------------------------------------
                sheet.Cells.InsertRows(3, 2); // insert after existing rows

                // Populate new tasks
                sheet.Cells["A4"].PutValue("Task 3");
                sheet.Cells["B4"].PutValue(DateTime.Today.AddDays(2));
                sheet.Cells["C4"].PutValue(DateTime.Today.AddDays(5));
                sheet.Cells["D4"].Formula = "=C4-B4";

                sheet.Cells["A5"].PutValue("Task 4");
                sheet.Cells["B5"].PutValue(DateTime.Today.AddDays(3));
                sheet.Cells["C5"].PutValue(DateTime.Today.AddDays(6));
                sheet.Cells["D5"].Formula = "=C5-B5";

                // -------------------------------------------------
                // 5. Resize the table to include the new rows
                // -------------------------------------------------
                table.Resize(0, 0, sheet.Cells.MaxDataRow, 3, true);

                // -------------------------------------------------
                // 6. Refresh chart data ranges to include new rows
                // -------------------------------------------------
                // Clear existing series and re‑add with updated ranges
                ganttChart.NSeries.Clear();

                string startRange = $"B2:B{sheet.Cells.MaxDataRow + 1}";
                string durationRange = $"D2:D{sheet.Cells.MaxDataRow + 1}";
                string categoryRange = $"A2:A{sheet.Cells.MaxDataRow + 1}";

                ganttChart.NSeries.Add(startRange, true);      // Start
                ganttChart.NSeries[0].IsColorVaried = false;   // keep invisible
                ganttChart.NSeries.Add(durationRange, true);   // Duration
                ganttChart.NSeries[1].IsColorVaried = true;

                ganttChart.NSeries.CategoryData = categoryRange;

                // -------------------------------------------------
                // 7. Recalculate the chart (ensures correct layout)
                // -------------------------------------------------
                ganttChart.Calculate();

                // -------------------------------------------------
                // 8. Save the workbook
                // -------------------------------------------------
                workbook.Save("UpdatedGanttChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}