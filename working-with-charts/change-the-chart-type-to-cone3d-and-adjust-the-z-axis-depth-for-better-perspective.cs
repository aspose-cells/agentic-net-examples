// Title: Convert a 3D Column Chart to Cone3D and Set DepthPercent in Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a 3‑D column chart, changes its type to Cone3D, sets DepthPercent to 250%, disables right‑angle axes, applies a 30° perspective, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Cone3D chart | DepthPercent | 3D chart perspective | chart type conversion | ChartType.Column3D | RightAngleAxes | Perspective property
// Common Searches: Aspose.Cells change chart type to Cone3D | Set DepthPercent for 3D chart Aspose.Cells C# | Adjust perspective of 3D chart in Aspose.Cells | Convert Column3D to Cone3D using Aspose.Cells | C# Aspose.Cells 3D chart depth and perspective
// Developer Intent: Switch an existing 3D column chart to a Cone3D chart and configure its depth and perspective settings using Aspose.Cells for .NET.
// Use Cases: Display sales trends with a Cone3D chart for enhanced visual emphasis. | Increase DepthPercent to 250% to make the 3D chart appear deeper. | Create a realistic 3D view by disabling RightAngleAxes and setting a custom perspective angle. | Automate generation of Excel reports that include customized 3D charts.
// AI Prompts: Generate C# code with Aspose.Cells that converts a Column3D chart to Cone3D, sets DepthPercent to 250, disables RightAngleAxes, and applies a 30-degree perspective. | Show how to adjust the Z‑axis depth and perspective of a 3D chart in Aspose.Cells for .NET. | Explain the effect of DepthPercent and Perspective properties on a Cone3D chart in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // This C# example creates a workbook, adds a 3‑D column chart, changes its type to Cone3D, sets DepthPercent to 250%, disables right‑angle axes, applies a 30° perspective, and saves the file.
    class Program
    {
        static void Main()
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
                worksheet.Cells["B2"].PutValue(1000);
                worksheet.Cells["B3"].PutValue(2000);
                worksheet.Cells["B4"].PutValue(3000);

                // Add a 3‑D column chart (initial type can be any 3‑D chart)
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Change the chart type to a supported 3‑D chart (Column3D used here)
                chart.Type = ChartType.Column3D;

                // Adjust the depth of the 3‑D chart for better perspective (percentage of chart width)
                chart.DepthPercent = 250; // 250% depth

                // Optional: fine‑tune perspective to enhance visual depth
                chart.RightAngleAxes = false;   // Enable perspective projection
                chart.Perspective = 30;         // Perspective angle (0‑100)

                // Define output file path
                string outputPath = "Cone3DChart.xlsx";

                // Ensure the directory exists (in case a relative path is used)
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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
