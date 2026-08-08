// Title: Create a Pivot Timeline with Logarithmic Date Axis, Custom Tick Intervals, and PNG Export using Aspose.Cells C#
// Description: This example builds a workbook, fills column A with monthly dates and column B with exponential values, creates a PivotTable, adds a Timeline linked to the Date field, inserts a line chart, configures the category axis for a base‑10 logarithmic date scale with major ticks every two months and minor ticks every month, sets axis titles, saves the workbook, and renders the chart to a PNG image without using System.Drawing.
// Keywords: Aspose.Cells timeline | logarithmic date axis | custom tick intervals | C# chart export PNG | pivot table timeline Aspose | .NET Excel chart rendering | log scale chart Aspose.Cells
// Common Searches: Aspose.Cells add timeline to pivot table | logarithmic date axis C# Aspose.Cells | custom month tick intervals chart Aspose | export Aspose.Cells chart to PNG | timeline control with log scale chart
// Developer Intent: Generate a workbook with a pivot‑linked timeline, a line chart that uses a logarithmic date axis and custom tick intervals, and save the chart as a PNG image.
// Use Cases: Show exponential growth over monthly periods while enabling interactive date filtering via a timeline. | Automate creation of PNG charts for dashboards or reports without relying on System.Drawing. | Develop financial or scientific models where a log‑scaled date axis clarifies large value ranges.
// AI Prompts: Modify the code to use a logarithmic base of 2 and set major ticks to every 3 months. | Provide a snippet that saves the timeline shape as a separate PNG file while preserving the workbook. | Explain how to position and size the timeline using pixel‑based properties instead of the Shape object.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineLogChart
{
    // This example builds a workbook, fills column A with monthly dates and column B with exponential values, creates a PivotTable, adds a Timeline linked to the Date field, inserts a line chart, configures the category axis for a base‑10 logarithmic date scale with major ticks every two months and minor ticks every month, sets axis titles, saves the workbook, and renders the chart to a PNG image without using System.Drawing.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate worksheet with sample date/value data
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Value");
                for (int i = 2; i <= 11; i++)
                {
                    // Dates spaced by one month
                    cells[$"A{i}"].PutValue(new DateTime(2023, i - 1, 1));
                    // Exponential values to illustrate logarithmic scaling
                    cells[$"B{i}"].PutValue(Math.Pow(2, i - 2));
                }

                // Create a PivotTable based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B11", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline linked to the PivotTable's Date field
                // Position the Timeline at row 15, column 1 (cell A15)
                int timelineIndex = sheet.Timelines.Add(pivot, 14, 0, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];
                timeline.Caption = "Date Timeline";
                // Adjust size using the Shape object (preferred over obsolete pixel properties)
                timeline.Shape.Width = 400;
                timeline.Shape.Height = 80;

                // Add a Line chart to visualize the same data
                int chartIndex = sheet.Charts.Add(ChartType.Line, 20, 0, 35, 15);
                Chart chart = sheet.Charts[chartIndex];
                // Set the data source for the chart
                chart.NSeries.Add("B2:B11", true);
                chart.NSeries[0].XValues = "A2:A11";
                chart.NSeries[0].Name = "Exponential Values";

                // Configure the category (X) axis to use a logarithmic scale for dates
                chart.CategoryAxis.IsLogarithmic = true;   // Enable logarithmic scaling
                chart.CategoryAxis.LogBase = 10;           // Base 10 logarithm
                // Customize tick intervals: major unit every 2 months, minor unit every 1 month
                chart.CategoryAxis.CategoryType = CategoryType.TimeScale;
                chart.CategoryAxis.BaseUnitScale = TimeUnit.Months;
                chart.CategoryAxis.MajorUnitScale = TimeUnit.Months;
                chart.CategoryAxis.MajorUnit = 2;          // Major tick every 2 months
                chart.CategoryAxis.MinorUnitScale = TimeUnit.Months;
                chart.CategoryAxis.MinorUnit = 1;          // Minor tick every month

                // Add titles for clarity
                chart.Title.Text = "Logarithmic Date Scale Chart";
                chart.CategoryAxis.Title.Text = "Date (Log Scale)";
                chart.ValueAxis.Title.Text = "Value";

                // Save the workbook (optional, for reference)
                workbook.Save("TimelineLogChart.xlsx");

                // Render the workbook (including the chart) to a PNG image
                // This avoids the need for System.Drawing dependencies
                workbook.Save("LogChart.png", SaveFormat.Png);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
