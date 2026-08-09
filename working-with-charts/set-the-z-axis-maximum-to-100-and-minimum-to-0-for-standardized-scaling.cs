// Title: Set Z‑Axis (Depth Axis) Minimum 0 and Maximum 100 for a 3‑D Column Chart in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a 3‑D clustered column chart, accesses the DepthAxis/ZAxis via reflection, disables automatic scaling, sets MinValue = 0 and MaxValue = 100, and saves the file as ChartZAxisScaling.xlsx.
// Keywords: Aspose.Cells | C# | .NET | 3D column chart | DepthAxis | ZAxis | axis minimum | axis maximum | fixed scaling | chart axis range | reflection
// Common Searches: Aspose.Cells set Z axis range C# | How to change depth axis limits in 3D chart Aspose.Cells | C# fix Z axis minimum and maximum Aspose.Cells | Access ZAxis property with reflection Aspose.Cells | Set custom scaling for 3D chart axis .NET
// Developer Intent: Set the Z‑axis (depth axis) of a 3‑D column chart to a fixed range of 0‑100.
// Use Cases: Standardize depth scaling across multiple reports for consistent visual comparison. | Enforce business rules that require chart depth values between 0 and 100. | Prepare workbooks for automated distribution where Excel must display a uniform Z‑axis. | Integrate chart generation into a .NET service that outputs Excel files with predefined axis limits.
// AI Prompts: Generate C# code using Aspose.Cells to set the Z‑axis minimum to 0 and maximum to 100 for a 3‑D column chart, handling both DepthAxis and ZAxis property names. | Explain how to disable automatic scaling and assign custom MinValue and MaxValue to a chart's depth axis in Aspose.Cells for .NET. | Show how to verify the Z‑axis limits after saving the workbook, e.g., by reading the axis properties or opening the file in Excel. | Provide a fallback approach if the chart type does not expose a ZAxis property.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a 3‑D clustered column chart, accesses the DepthAxis/ZAxis via reflection, disables automatic scaling, sets MinValue = 0 and MaxValue = 100, and saves the file as ChartZAxisScaling.xlsx.
    public class SetZAxisScalingDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(30);
            worksheet.Cells["B4"].PutValue(50);

            // Add a 3‑D column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column3DClustered, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the depth (Z) axis via reflection (DepthAxis or ZAxis depending on version)
            Axis depthAxis = null;
            var depthProp = chart.GetType().GetProperty("DepthAxis");
            if (depthProp != null)
            {
                depthAxis = depthProp.GetValue(chart) as Axis;
            }
            else
            {
                var zProp = chart.GetType().GetProperty("ZAxis");
                if (zProp != null)
                {
                    depthAxis = zProp.GetValue(chart) as Axis;
                }
            }

            // Apply fixed scaling if the axis was found
            if (depthAxis != null)
            {
                depthAxis.IsAutomaticMinValue = false; // Disable automatic minimum
                depthAxis.IsAutomaticMaxValue = false; // Disable automatic maximum
                depthAxis.MinValue = 0;                // Set minimum to 0
                depthAxis.MaxValue = 100;              // Set maximum to 100
            }
            else
            {
                Console.WriteLine("Depth (Z) axis not available for this chart type.");
            }

            // Save the workbook
            string outputPath = "ChartZAxisScaling.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
