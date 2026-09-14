// Title: Create a 3‑D Cone chart (fallback to Column3D) and configure Z‑axis depth and perspective with Aspose.Cells for .NET
// AI Prompts: Generate a new workbook, add sample data, insert a 3‑D chart, set its type to Cone3D (or Column3D if Cone3D is not exposed), and apply DepthPercent, Perspective, RightAngleAxes, RotationAngle, and Elevation using Aspose.Cells in C#. | Replace an existing 3‑D column chart with a Cone3D chart and modify its Z‑axis depth and perspective settings to improve visual depth in a .NET application.
// Common Searches: how to create a cone 3d chart with Aspose.Cells C# when Cone3D type is unavailable | set depth percent and perspective for a 3d column chart using Aspose.Cells .NET | adjust Z axis depth of a 3d chart in Aspose.Cells C# example | change chart type to Cone3D and tweak rotation angle in Aspose.Cells for .NET | Aspose.Cells example for configuring 3d chart perspective and elevation
// Tags: Aspose.Cells create cone3d chart C# | set chart depthpercent Aspose.Cells | configure 3d chart perspective Aspose.Cells | replace column3d with cone3d Aspose.Cells | adjust rotationangle elevation 3d chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example builds a workbook, fills cells A1:B4 with category and sales data, adds a 3‑D column chart as a placeholder for Cone3D, binds the series to the data range, sets DepthPercent to 250, Perspective to 40, disables RightAngleAxes, and fine‑tunes RotationAngle and Elevation before saving the file as Cone3DChartDemo.xlsx.
    public class Cone3DChartDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");

                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(1200);
                worksheet.Cells["B3"].PutValue(1500);
                worksheet.Cells["B4"].PutValue(1800);

                // Add a 3‑D column chart (Aspose.Cells version used does not expose Cone3D, so Column3D is used)
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // If Cone3D becomes available, replace the line below with:
                // chart.Type = ChartType.Cone3D;
                // For now keep the 3‑D column type.
                chart.Type = ChartType.Column3D;

                // Adjust depth (Z‑axis) for better perspective
                chart.DepthPercent = 250;          // 250% depth (range 20‑2000)
                chart.Perspective = 40;            // Perspective angle (0‑100)
                chart.RightAngleAxes = false;      // Enable perspective projection

                // Optional: tweak rotation/elevation for visual effect
                chart.RotationAngle = 30;
                chart.Elevation = 20;

                // Determine output path and ensure directory exists
                string outputPath = "Cone3DChartDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Cone3D chart demo: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            Cone3DChartDemo.Run();
        }
    }
}
