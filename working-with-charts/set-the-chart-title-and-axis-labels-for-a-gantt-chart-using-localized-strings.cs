// Title: Set Localized Chart Title and Axis Labels for a Gantt Chart with Aspose.Cells for .NET
// Description: Shows how to build a workbook, insert task data, create a bar‑type Gantt chart, and apply Chinese text to the chart title and value‑axis label using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# chart title | localized chart labels | Gantt chart | Chinese chart title | Excel chart localization | Bar chart as Gantt | set axis title | Aspose.Cells .NET
// Common Searches: Aspose.Cells set Chinese chart title | how to localize axis label in Aspose.Cells C# | create Gantt chart with localized headings | C# Aspose.Cells chart title language | set value axis title for bar chart Aspose
// Developer Intent: Add or modify chart title and axis labels in a chosen language for a Gantt‑style chart using Aspose.Cells.
// Use Cases: Produce project‑schedule workbooks with Chinese headings for regional stakeholders. | Automate Excel reports that display language‑specific chart titles across multiple locales. | Generate bar‑type Gantt charts that meet multilingual documentation standards.
// AI Prompts: Generate C# code with Aspose.Cells that reads a resource file and assigns localized strings to a Gantt chart's title and axis labels. | Demonstrate how to set non‑English text for the Title and ValueAxis.Title of a bar chart used as a Gantt chart in Aspose.Cells. | Provide an example that toggles visibility and applies localized text to chart titles and axis labels in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttChartExample
{
    // Shows how to build a workbook, insert task data, create a bar‑type Gantt chart, and apply Chinese text to the chart title and value‑axis label using Aspose.Cells in C#.
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
                // Prepare sample data for a Gantt chart
                // -------------------------------------------------
                // Columns: Task, Start Date, End Date
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("End");

                sheet.Cells["A2"].PutValue("Planning");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 10));

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 11));
                sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 20));

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 21));
                sheet.Cells["C4"].PutValue(new DateTime(2023, 3, 5));

                // -------------------------------------------------
                // Add a chart (using Bar chart to simulate Gantt)
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 25, 15);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Set data source: Category (Task) and Start/End dates
                ganttChart.NSeries.Add("B2:C4", true);
                ganttChart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Set localized titles directly
                // -------------------------------------------------
                ganttChart.Title.Text = "项目进度甘特图"; // "Project Schedule Gantt Chart" in Chinese
                ganttChart.Title.IsVisible = true;

                // Value axis (horizontal) represents the time scale
                ganttChart.ValueAxis.Title.Text = "时间轴"; // "Time Axis" in Chinese
                ganttChart.ValueAxis.Title.IsVisible = true;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "GanttChartLocalized.xlsx";
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
