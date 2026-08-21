// Title: Aspose.Cells for .NET: Convert a 3D Column Chart to a 3D Cone and Enhance Z‑Axis Depth
// Description: This example creates a workbook, adds sample sales data, inserts a 3‑D column chart, switches the chart type to a 3‑D cone, sets DepthPercent to 250, adjusts Perspective to 40, disables right‑angle axes, and fine‑tunes RotationAngle and Elevation before saving the file.
// Keywords: Aspose.Cells C# 3D cone chart | change chart type to cone Aspose.Cells | DepthPercent property Aspose.Cells | chart perspective Aspose.Cells .NET | right angle axes false Aspose.Cells | rotation angle elevation chart | 3D chart customization .NET | GitHub Aspose.Cells examples
// Common Searches: how to change a 3d column chart to a cone chart using Aspose.Cells | increase depth percent of a 3d chart in C# | set perspective and rotation for Aspose.Cells 3D charts | disable right‑angle axes for better 3D view Aspose.Cells | Aspose.Cells example for cone3d chart
// Developer Intent: Transform an existing 3‑D column chart into a 3‑D cone chart and improve its visual depth by configuring depth, perspective, rotation, and elevation settings.
// Use Cases: Generate sales reports with a 3‑D cone chart that stands out in presentations. | Create financial dashboards where cone charts emphasize category differences with custom 3‑D angles. | Automate chart styling across multiple worksheets to maintain a consistent 3‑D appearance.
// AI Prompts: Show C# code that converts a 3D column chart to a cone chart and sets DepthPercent, Perspective, RotationAngle, and Elevation using Aspose.Cells. | Explain the valid ranges for DepthPercent and Perspective properties in Aspose.Cells and demonstrate their application. | Provide a step‑by‑step guide to disable right‑angle axes and adjust rotation for a 3D cone chart in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample sales data, inserts a 3‑D column chart, switches the chart type to a 3‑D cone, sets DepthPercent to 250, adjusts Perspective to 40, disables right‑angle axes, and fine‑tunes RotationAngle and Elevation before saving the file.
    public class ChangeChartToCone3D
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["B4"].PutValue(3000);

            // Add an initial 3‑D column chart (will be changed to Cone later)
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the chart type to Cone (3‑D cone not available in older versions)
            chart.Type = ChartType.Cone;

            // Adjust Z‑axis depth for a stronger 3‑D effect (valid range 20‑2000)
            chart.DepthPercent = 250; // 250% depth

            // Set perspective to improve visual depth (0‑100)
            chart.Perspective = 40;
            chart.RightAngleAxes = false; // Enable perspective projection

            // Optional: tweak rotation and elevation for better view
            chart.RotationAngle = 20;
            chart.Elevation = 15;

            // Save the workbook with the modified chart
            string outputPath = "ChartCone3D_WithDepth.xlsx";
            workbook.Save(outputPath);
        }
    }
}
