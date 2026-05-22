using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartCustomPaletteDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Prepare sample data: Task, Start Date, Duration, Priority
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");
                sheet.Cells["D1"].PutValue("Priority");

                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["D2"].PutValue("High");

                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 12));
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["D3"].PutValue("Medium");

                sheet.Cells["A4"].PutValue("Testing");
                sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 5));
                sheet.Cells["C4"].PutValue(8);
                sheet.Cells["D4"].PutValue("Low");

                // -------------------------------------------------
                // Define custom colors for priorities using ChangePalette (rule)
                // Index 0 -> High (Red), Index 1 -> Medium (Orange), Index 2 -> Low (Green)
                // -------------------------------------------------
                Color highColor = Color.Red;
                Color mediumColor = Color.Orange;
                Color lowColor = Color.Green;

                workbook.ChangePalette(highColor, 0);
                workbook.ChangePalette(mediumColor, 1);
                workbook.ChangePalette(lowColor, 2);

                // -------------------------------------------------
                // Add a Gantt‑like chart (implemented as a stacked bar chart)
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 25, 10);
                Chart ganttChart = sheet.Charts[chartIndex];

                // Set the data range for the series (Start and Duration)
                // Gantt chart expects start dates in first column and durations in second column
                ganttChart.NSeries.Add("B2:C4", true);
                ganttChart.NSeries.CategoryData = "A2:A4";

                // Ensure each bar can have its own color
                ganttChart.NSeries.IsColorVaried = false;

                // -------------------------------------------------
                // Apply colors to each task bar based on priority
                // -------------------------------------------------
                for (int i = 0; i < ganttChart.NSeries[0].Points.Count; i++)
                {
                    // Read priority from column D
                    string priority = sheet.Cells[i + 2, 3].StringValue.Trim().ToLower();

                    // Choose palette index
                    int paletteIndex = priority switch
                    {
                        "high" => 0,
                        "medium" => 1,
                        "low" => 2,
                        _ => 0 // default to high if unknown
                    };

                    // Assign the corresponding palette color to the point (bar)
                    ganttChart.NSeries[0].Points[i].Area.ForegroundColor = workbook.Colors[paletteIndex];
                }

                // -------------------------------------------------
                // Save the workbook (lifecycle rule)
                // -------------------------------------------------
                string outputPath = "GanttChartCustomPalette.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}