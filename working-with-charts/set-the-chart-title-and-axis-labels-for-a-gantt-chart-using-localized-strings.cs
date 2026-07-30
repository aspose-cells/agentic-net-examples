// Title: C# – Set Chinese Chart Title and Axis Labels for a Gantt Chart using Aspose.Cells
// Description: Demonstrates how to create a workbook, add a stacked‑bar chart that mimics a Gantt chart, and assign Chinese strings to the chart title, category axis, and value axis with Aspose.Cells for .NET, then save the file as an Excel workbook.
// Keywords: Aspose.Cells C# | Gantt chart localization | set chart title Aspose.Cells | axis labels Chinese | Unicode chart text | stacked bar Gantt Aspose | Excel chart title .NET | multilingual Excel reports | Aspose.Cells example
// Common Searches: Aspose.Cells set chart title C# | how to localize axis labels in Aspose.Cells | Chinese chart title Aspose.Cells .NET | create Gantt chart with Aspose.Cells | Aspose.Cells chart localization tutorial
// Developer Intent: Apply Chinese (Unicode) text to the chart title and both axis titles of a Gantt‑style chart using Aspose.Cells for .NET.
// Use Cases: Generate project‑schedule workbooks with Chinese‑language chart headings for teams in China. | Produce multilingual Excel dashboards where chart titles and axis labels must appear in the target locale. | Automate creation of localized Gantt charts for international project management reporting.
// AI Prompts: Show C# code that sets a Unicode chart title and axis labels for a Gantt chart with Aspose.Cells. | Explain how to localize chart titles and axis captions in a stacked‑bar chart, including handling right‑to‑left languages. | Provide steps to read an existing Excel file, modify its chart title and axis text, and save the changes using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttChartLocalization
{
    // Demonstrates how to create a workbook, add a stacked‑bar chart that mimics a Gantt chart, and assign Chinese strings to the chart title, category axis, and value axis with Aspose.Cells for .NET, then save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate minimal data for a Gantt‑like chart (Task, Start, Duration)
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 12));
                sheet.Cells["C3"].PutValue(20);

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 5));
                sheet.Cells["C4"].PutValue(5);

                // Add a stacked bar chart (used to emulate a Gantt chart)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Bind data: Duration as series, Task names as categories, Start dates as X values
                ganttChart.NSeries.Add("C2:C4", true);
                ganttChart.NSeries.CategoryData = "A2:A4";
                ganttChart.NSeries[0].XValues = "B2:B4";

                // Localized titles (Chinese)
                string chartTitle = "项目进度甘特图"; // "Project Schedule Gantt Chart"
                string axisTitle = "时间轴";          // "Time Axis"

                ganttChart.Title.Text = chartTitle;
                ganttChart.CategoryAxis.Title.Text = "任务"; // "Task"
                ganttChart.ValueAxis.Title.Text = axisTitle;

                // Save the workbook
                string outputPath = "LocalizedGanttChart.xlsx";
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
