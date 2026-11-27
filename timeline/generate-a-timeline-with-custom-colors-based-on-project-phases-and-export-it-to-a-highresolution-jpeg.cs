// Title: Generate a project timeline chart with phase‑specific colors and export it as a 300 DPI JPEG using Aspose.Cells for .NET
// Common Searches: how to color chart bars based on a worksheet column in Aspose.Cells C# | export Aspose.Cells chart to high‑resolution JPEG image in .NET | create a Gantt‑like timeline using a column chart in Aspose.Cells | set custom DPI when saving an Aspose.Cells chart as JPEG | apply phase colors to timeline bars from Excel data with Aspose.Cells
// Tags: data point color mapping Aspose.Cells | high DPI JPEG export Aspose.Cells chart | using column chart for project timeline Aspose.Cells | phase based bar coloring C# Aspose.Cells | embed chart in workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace TimelineExportDemo
{
    // The sample program loads or creates an Excel workbook containing tasks, start/end dates, and phase information, builds a column chart to represent the timeline, assigns custom colors to each bar based on the Phase column, and then exports the chart as a 300 DPI JPEG image while also saving the workbook with the chart embedded.
    class Program
    {
        static void Main(string[] args)
        {
            const string workbookPath = "ProjectTimeline.xlsx";
            const string outputImagePath = "ProjectTimeline.jpg";

            Workbook workbook = null;
            Worksheet worksheet = null;

            // Load existing workbook or create a sample one
            if (File.Exists(workbookPath))
            {
                try
                {
                    workbook = new Workbook(workbookPath);
                    worksheet = workbook.Worksheets[0];
                    Console.WriteLine($"Loaded workbook from \"{workbookPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook \"{workbookPath}\": {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Workbook \"{workbookPath}\" not found. Creating a sample workbook.");
                workbook = new Workbook();
                worksheet = workbook.Worksheets[0];

                // Header row
                worksheet.Cells["A1"].PutValue("Task");
                worksheet.Cells["B1"].PutValue("StartDate");
                worksheet.Cells["C1"].PutValue("EndDate");
                worksheet.Cells["D1"].PutValue("Phase");

                // Sample data
                var sampleData = new[]
                {
                    new {Task = "Requirement Analysis", Start = new DateTime(2023, 1, 1), End = new DateTime(2023, 1, 15), Phase = "Planning"},
                    new {Task = "Design",               Start = new DateTime(2023, 1, 16), End = new DateTime(2023, 2, 5),  Phase = "Design"},
                    new {Task = "Implementation",       Start = new DateTime(2023, 2, 6),  End = new DateTime(2023, 4, 30), Phase = "Execution"},
                    new {Task = "Testing",              Start = new DateTime(2023, 5, 1),  End = new DateTime(2023, 5, 31), Phase = "Testing"},
                    new {Task = "Deployment",           Start = new DateTime(2023, 6, 1),  End = new DateTime(2023, 6, 15), Phase = "Release"}
                };

                for (int i = 0; i < sampleData.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2
                    worksheet.Cells[$"A{row}"].PutValue(sampleData[i].Task);
                    worksheet.Cells[$"B{row}"].PutValue(sampleData[i].Start);
                    worksheet.Cells[$"C{row}"].PutValue(sampleData[i].End);
                    worksheet.Cells[$"D{row}"].PutValue(sampleData[i].Phase);
                }

                // Save the sample workbook for future runs
                try
                {
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Sample workbook saved as \"{workbookPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save sample workbook: {ex.Message}");
                }
            }

            // Create a chart (used as a timeline representation)
            Chart chart;
            try
            {
                // Using Column chart as a substitute for Gantt (Aspose.Cells does not expose Gantt)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 30, 10);
                chart = worksheet.Charts[chartIndex];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create chart: {ex.Message}");
                return;
            }

            // Add series: start dates and end dates
            chart.NSeries.Add("B2:B6", true); // Start dates
            chart.NSeries.Add("C2:C6", true); // End dates

            // Set category (task names) for both series
            chart.NSeries.CategoryData = "A2:A6";

            // Optional title
            chart.Title.Text = "Project Timeline";

            // Map phases to colors
            var phaseColors = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
            {
                { "Planning",  Color.LightBlue },
                { "Design",    Color.Orange    },
                { "Execution", Color.LightGreen},
                { "Testing",   Color.Violet    },
                { "Release",   Color.Gold      }
            };

            // Apply custom colors to each data point based on the Phase column (D)
            try
            {
                for (int i = 0; i < chart.NSeries[1].Points.Count; i++)
                {
                    int dataRow = i + 2;
                    string phase = worksheet.Cells[$"D{dataRow}"].StringValue;
                    Color clr = phaseColors.TryGetValue(phase, out Color c) ? c : Color.Gray;
                    chart.NSeries[1].Points[i].Area.ForegroundColor = clr;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to apply custom colors: {ex.Message}");
            }

            // High‑resolution image options
            var imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Export chart to JPEG
            try
            {
                chart.ToImage(outputImagePath, imgOptions);
                Console.WriteLine($"Timeline chart exported successfully to \"{outputImagePath}\" with 300 DPI resolution.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart to image: {ex.Message}");
                return;
            }

            // Save workbook with the chart embedded
            const string savedWorkbookPath = "ProjectTimeline_WithChart.xlsx";
            try
            {
                workbook.Save(savedWorkbookPath);
                Console.WriteLine($"Workbook with chart saved as \"{savedWorkbookPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
