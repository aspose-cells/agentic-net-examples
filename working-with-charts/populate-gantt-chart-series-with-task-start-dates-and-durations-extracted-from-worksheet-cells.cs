// Title: Create a Gantt chart in C# with Aspose.Cells by reading task start dates and durations from worksheet cells
// Common Searches: C# Aspose.Cells generate Gantt chart from start date and duration columns | how to hide the start‑date series in a stacked bar Gantt chart using Aspose.Cells | populate chart series with date values in Aspose.Cells C# example | Aspose.Cells create project schedule Gantt chart with stacked bar chart | set date number format for cells when building a Gantt chart with Aspose.Cells | add title and remove legend from a Gantt chart using Aspose.Cells C#
// Tags: Aspose.Cells create stacked bar Gantt chart | Aspose.Cells populate series from worksheet range | Aspose.Cells hide series fill and border | Aspose.Cells apply date number format to cells | Aspose.Cells add chart title hide legend | Aspose.Cells save workbook as XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace GanttChartDemo
{
    // The example creates a new workbook, writes task names, start dates (formatted as dates) and durations into columns A‑C, adds a stacked bar chart, defines two series (start dates hidden, durations visible), sets the category axis to the task names, removes fill and border from the start‑date series, adds a chart title, hides the legend, and saves the result as GanttChartDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample task data.
                //    Column A : Task Name
                //    Column B : Start Date
                //    Column C : Duration (in days)
                // -----------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                var tasks = new[]
                {
                    new { Name = "Define Scope",   Start = new DateTime(2024, 1, 2), Duration = 5 },
                    new { Name = "Design",         Start = new DateTime(2024, 1, 8), Duration = 8 },
                    new { Name = "Implementation", Start = new DateTime(2024, 1, 17), Duration = 12 },
                    new { Name = "Testing",        Start = new DateTime(2024, 2, 2), Duration = 6 },
                    new { Name = "Deployment",     Start = new DateTime(2024, 2, 10), Duration = 3 }
                };

                for (int i = 0; i < tasks.Length; i++)
                {
                    int row = i + 2; // data starts at row 2
                    sheet.Cells[$"A{row}"].PutValue(tasks[i].Name);
                    sheet.Cells[$"B{row}"].PutValue(tasks[i].Start);
                    sheet.Cells[$"C{row}"].PutValue(tasks[i].Duration);

                    // Apply date number format to the start column
                    Style dateStyle = workbook.CreateStyle();
                    dateStyle.Number = 14; // m/d/yy
                    sheet.Cells[$"B{row}"].SetStyle(dateStyle);
                }

                // -----------------------------------------------------------------
                // 3. Add a stacked bar chart (the typical Gantt representation).
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // -----------------------------------------------------------------
                // 4. Add two series:
                //    - Series 0 : Start dates (will be hidden)
                //    - Series 1 : Durations (visible bars)
                // -----------------------------------------------------------------
                chart.NSeries.Add("B2:B6", false); // Start dates
                chart.NSeries.Add("C2:C6", false); // Durations
                chart.NSeries.CategoryData = "A2:A6";

                // -----------------------------------------------------------------
                // 5. Hide the start‑date series so that only the duration bars appear.
                // -----------------------------------------------------------------
                Series startSeries = chart.NSeries[0];
                startSeries.Area.FillFormat.Pattern = FillPattern.None; // No fill
                startSeries.Border.IsVisible = false;                    // No border

                // -----------------------------------------------------------------
                // 6. Optional formatting (date axis). This step is skipped if the API is unavailable.
                // -----------------------------------------------------------------
                // Note: Axis.NumberFormat may not be supported in some versions; therefore it is omitted.

                // -----------------------------------------------------------------
                // 7. Add a title and hide the legend.
                // -----------------------------------------------------------------
                chart.Title.Text = "Project Schedule (Gantt Chart)";
                chart.ShowLegend = false;

                // -----------------------------------------------------------------
                // 8. Save the workbook.
                // -----------------------------------------------------------------
                string outFile = "GanttChartDemo.xlsx";
                try
                {
                    workbook.Save(outFile, SaveFormat.Xlsx);
                    Console.WriteLine($"Gantt chart saved to '{outFile}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
