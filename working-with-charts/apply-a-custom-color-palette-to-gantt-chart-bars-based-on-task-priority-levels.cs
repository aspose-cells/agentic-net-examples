// Title: Color Gantt Chart Bars by Priority with a Custom Palette in Aspose.Cells (C#)
// Description: Demonstrates how to create a Gantt‑style stacked bar chart in a new workbook, define a three‑color palette for priority levels, hide the start‑date series, and assign each task bar a color from the palette based on its priority before saving the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET | Gantt chart | custom palette | stacked bar chart | chart series color | priority based coloring | ChangePalette | IsColorVaried | conditional chart formatting
// Common Searches: Aspose.Cells set custom colors for Gantt chart bars | C# change workbook palette for chart series | make first series transparent stacked bar chart Aspose.Cells | apply per‑point colors based on data Aspose.Cells | color code tasks by priority in Excel using Aspose
// Developer Intent: Generate a Gantt‑style chart where each task bar is colored according to its priority using a custom workbook palette.
// Use Cases: Define red, orange, and green palette entries for high, medium, and low priority tasks. | Create a stacked bar chart that mimics a Gantt chart and hide the start‑date series. | Read priority values from a worksheet column and apply the matching palette color to each duration bar. | Save the configured workbook as an XLSX file for distribution or further analysis.
// AI Prompts: Write C# code that reads a priority column and colors each point in an Aspose.Cells stacked bar chart using a custom palette. | Explain how to modify the workbook palette with ChangePalette and retrieve colors for chart points in Aspose.Cells for .NET. | Show steps to make the first series of a stacked bar chart transparent while varying the second series colors based on data.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartCustomPaletteDemo
{
    // Demonstrates how to create a Gantt‑style stacked bar chart in a new workbook, define a three‑color palette for priority levels, hide the start‑date series, and assign each task bar a color from the palette based on its priority before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Define custom colors for priority levels and add them to the palette
                // -------------------------------------------------
                // Priority 1 -> Red (palette index 0)
                // Priority 2 -> Orange (palette index 1)
                // Priority 3 -> Green (palette index 2)
                workbook.ChangePalette(Color.FromArgb(255, 200, 0, 0), 0); // Red
                workbook.ChangePalette(Color.FromArgb(255, 255, 165, 0), 1); // Orange
                workbook.ChangePalette(Color.FromArgb(255, 0, 128, 0), 2); // Green

                // -------------------------------------------------
                // Populate sample data: Task, Start Date (as number), Duration, Priority
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Start");
                sheet.Cells["C1"].PutValue("Duration");
                sheet.Cells["D1"].PutValue("Priority");

                // Sample rows
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["B2"].PutValue(1);   // Start day
                sheet.Cells["C2"].PutValue(5);   // Duration
                sheet.Cells["D2"].PutValue(1);   // Priority 1

                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["B3"].PutValue(3);
                sheet.Cells["C3"].PutValue(4);
                sheet.Cells["D3"].PutValue(2);   // Priority 2

                sheet.Cells["A4"].PutValue("Task 4");
                sheet.Cells["B4"].PutValue(2);
                sheet.Cells["C4"].PutValue(6);
                sheet.Cells["D4"].PutValue(3);   // Priority 3

                // -------------------------------------------------
                // Add a Gantt‑like chart (stacked bar chart where the first series is invisible)
                // -------------------------------------------------
                // Use BarStacked as the compatible enum for stacked bar charts
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 10);
                Chart ganttChart = sheet.Charts[chartIndex];

                // First series – Start (will be made transparent)
                // Second series – Duration (visible bars)
                ganttChart.NSeries.Add("B2:C4", true); // B column = start, C column = duration
                ganttChart.NSeries.CategoryData = "A2:A4";

                // Hide the start series by making its points fully transparent
                Series startSeries = ganttChart.NSeries[0];
                for (int i = 0; i < startSeries.Points.Count; i++)
                {
                    startSeries.Points[i].Area.ForegroundColor = Color.Transparent;
                }

                // -------------------------------------------------
                // Apply custom colors to each duration bar based on priority
                // -------------------------------------------------
                Series durationSeries = ganttChart.NSeries[1];

                for (int i = 0; i < durationSeries.Points.Count; i++)
                {
                    // Read priority from column D (zero‑based index 3)
                    int priority = Convert.ToInt32(sheet.Cells[i + 2, 3].Value);
                    int paletteIndex = Math.Max(0, priority - 1); // safeguard index

                    // Retrieve the custom color from the workbook palette
                    Color barColor = workbook.Colors[paletteIndex];

                    // Apply the color to the point (bar)
                    durationSeries.Points[i].Area.ForegroundColor = barColor;
                    durationSeries.Points[i].Area.Formatting = FormattingType.Custom;
                }

                // Allow each point to have its own color
                durationSeries.IsColorVaried = true;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "GanttChartCustomPalette.xlsx";

                try
                {
                    // Ensure the directory exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
