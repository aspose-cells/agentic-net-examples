// Title: Color Gantt Chart Bars by Priority with Aspose.Cells for .NET
// Description: Creates an Excel workbook, adds task data with a priority column, builds a stacked‑bar Gantt chart, defines a custom palette entry, and assigns red, orange, or green colors to each bar based on its priority before saving the file.
// Keywords: Aspose.Cells | C# Gantt chart | custom chart colors | priority based coloring | stacked bar chart formatting | Workbook.ChangePalette | chart point foreground color | .NET Excel automation
// Common Searches: Aspose.Cells set Gantt bar color by priority | C# change palette for Excel chart Aspose | conditional colors for stacked bar series Aspose.Cells | how to color chart points in Aspose.Cells .NET | custom orange palette index Aspose.Cells
// Developer Intent: Apply priority‑specific colors to the bars of a Gantt‑style stacked bar chart using Aspose.Cells for .NET.
// Use Cases: Visual project schedules where high‑priority tasks appear in red, medium in orange, and low in green. | Maintain brand consistency by adding a custom palette entry and reusing it across multiple chart elements. | Generate automated Excel reports with pre‑formatted Gantt charts that require no manual styling.
// AI Prompts: Generate C# code that reads a priority column and colors each point of a stacked bar Gantt chart red, orange, or green with Aspose.Cells. | Show how to add a custom color to the workbook palette and apply it to chart points in a Gantt chart using Aspose.Cells for .NET. | Provide an example that saves an Excel file containing a priority‑colored Gantt chart and explain how to extend it for additional priority levels.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates an Excel workbook, adds task data with a priority column, builds a stacked‑bar Gantt chart, defines a custom palette entry, and assigns red, orange, or green colors to each bar based on its priority before saving the file.
class GanttChartCustomPalette
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate task data: Task name, Start date, Duration, Priority
            // -------------------------------------------------
            ws.Cells["A1"].PutValue("Task");
            ws.Cells["B1"].PutValue("Start");
            ws.Cells["C1"].PutValue("Duration");
            ws.Cells["D1"].PutValue("Priority");

            ws.Cells["A2"].PutValue("Task 1");
            ws.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            ws.Cells["C2"].PutValue(5);
            ws.Cells["D2"].PutValue("High");

            ws.Cells["A3"].PutValue("Task 2");
            ws.Cells["B3"].PutValue(new DateTime(2023, 1, 3));
            ws.Cells["C3"].PutValue(8);
            ws.Cells["D3"].PutValue("Medium");

            ws.Cells["A4"].PutValue("Task 3");
            ws.Cells["B4"].PutValue(new DateTime(2023, 1, 5));
            ws.Cells["C4"].PutValue(4);
            ws.Cells["D4"].PutValue("Low");

            // -------------------------------------------------
            // Add a Gantt‑style chart (implemented as a stacked bar chart)
            // -------------------------------------------------
            int chartIdx = ws.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range: start date + duration as series, tasks as categories
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Define a custom color in the workbook palette (optional)
            // -------------------------------------------------
            workbook.ChangePalette(Color.Orange, 55); // custom orange for medium priority

            // -------------------------------------------------
            // Apply colors to each Gantt bar based on priority
            // -------------------------------------------------
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                // Read priority from column D (zero‑based column index 3)
                string priority = ws.Cells[i + 2, 3].StringValue;
                Color barColor;

                switch (priority)
                {
                    case "High":
                        barColor = Color.Red;
                        break;
                    case "Medium":
                        barColor = Color.FromArgb(255, 165, 0); // custom orange
                        break;
                    case "Low":
                        barColor = Color.Green;
                        break;
                    default:
                        barColor = Color.Gray;
                        break;
                }

                // Apply the color to the point (the Gantt bar)
                chart.NSeries[0].Points[i].Area.ForegroundColor = barColor;
                chart.NSeries[0].Points[i].Area.Formatting = FormattingType.Custom;
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("GanttCustomPalette.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
