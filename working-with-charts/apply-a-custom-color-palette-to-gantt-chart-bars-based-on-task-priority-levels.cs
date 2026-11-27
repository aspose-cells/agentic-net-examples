// Title: Apply a priority‑based custom color palette to Gantt chart bars using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set bar color in stacked chart based on cell value | Create Gantt chart with custom priority colors using Aspose.Cells .NET | Change workbook palette and assign point colors in Aspose.Cells C# example | Conditional formatting of Gantt chart bars by priority in Aspose.Cells | How to color code Gantt chart tasks in Excel with Aspose.Cells API
// Tags: priority based bar coloring Aspose.Cells | custom palette stacked bar chart .NET | set chart point foreground color C# | Gantt chart color coding Aspose.Cells | change workbook palette Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartCustomPalette
{
    // The example creates a workbook, defines red, orange, and green colors for high, medium, and low priorities, updates the workbook palette, builds a stacked bar chart that mimics a Gantt schedule, and programmatically assigns each task bar a color according to its priority before saving the file as GanttChartWithCustomPalette.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Tasks";

                // -------------------------------------------------
                // 1. Prepare sample data for a Gantt chart
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start Date");
                sheet.Cells["C1"].PutValue("Duration (days)");
                sheet.Cells["D1"].PutValue("Priority"); // High, Medium, Low

                // Sample tasks
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2024, 1, 2));
                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["D2"].PutValue("High");

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2024, 1, 8));
                sheet.Cells["C3"].PutValue(10);
                sheet.Cells["D3"].PutValue("Medium");

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2024, 1, 20));
                sheet.Cells["C4"].PutValue(4);
                sheet.Cells["D4"].PutValue("Low");

                sheet.Cells["A5"].PutValue("Deployment");
                sheet.Cells["B5"].PutValue(new DateTime(2024, 1, 25));
                sheet.Cells["C5"].PutValue(2);
                sheet.Cells["D5"].PutValue("High");

                // -------------------------------------------------
                // 2. Define custom colors for priorities
                // -------------------------------------------------
                Color highPriorityColor   = Color.FromArgb(255, 0, 0);      // Red
                Color mediumPriorityColor = Color.FromArgb(255, 165, 0);    // Orange
                Color lowPriorityColor    = Color.FromArgb(0, 128, 0);      // Green

                // Change the workbook palette (indices 0‑2 are arbitrary but must be within 0‑55)
                workbook.ChangePalette(highPriorityColor,   0);
                workbook.ChangePalette(mediumPriorityColor, 1);
                workbook.ChangePalette(lowPriorityColor,    2);

                // -------------------------------------------------
                // 3. Create a Gantt‑like chart using a stacked bar chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 7, 0, 25, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Series 0 – Start dates (will be made invisible)
                chart.NSeries.Add("B2:B5", true);
                // Series 1 – Durations (visible)
                chart.NSeries.Add("C2:C5", true);

                // Set categories (task names)
                chart.NSeries.CategoryData = "A2:A5";

                // Chart title
                chart.Title.Text = "Project Schedule";

                // Make the start‑date series invisible by setting its color to transparent
                var startSeries = chart.NSeries[0];
                startSeries.Area.ForegroundColor = Color.Transparent;

                // -------------------------------------------------
                // 4. Apply custom colors to each bar based on priority
                // -------------------------------------------------
                var durationSeries = chart.NSeries[1];

                // Loop through each task row and set point color
                for (int i = 0; i < sheet.Cells.MaxDataRow - 1; i++) // data rows start at row 2 (index 1)
                {
                    int dataRow = i + 2; // Excel row number (2‑based)
                    string priority = sheet.Cells[$"D{dataRow}"].StringValue.Trim().ToLowerInvariant();

                    Color pointColor = priority switch
                    {
                        "high"   => highPriorityColor,
                        "medium" => mediumPriorityColor,
                        "low"    => lowPriorityColor,
                        _        => Color.Gray
                    };

                    // Apply the color to the corresponding point
                    ChartPoint point = durationSeries.Points[i];
                    point.Area.ForegroundColor = pointColor;
                }

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                string outputPath = "GanttChartWithCustomPalette.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
