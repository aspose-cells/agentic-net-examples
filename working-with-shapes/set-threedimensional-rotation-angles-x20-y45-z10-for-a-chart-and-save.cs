// Title: Aspose.Cells C# Example: Set 3‑D Chart Rotation (X=20°, Y=45°, Z=10°) and Save Workbook
// Description: A ready‑to‑run C# snippet that creates a workbook with Aspose.Cells, adds a 3‑D column chart, fills it with sample data, applies Elevation = 20°, Perspective = 45, RotationAngle = 10°, and saves the result as Chart3DRotationDemo.xlsx. Ideal for .NET developers needing precise chart orientation.
// Keywords: Aspose.Cells 3D chart rotation | C# set chart elevation | Aspose.Cells perspective property | RotationAngle Aspose.Cells | save workbook C# Aspose | Excel 3D column chart example | Aspose.Cells GitHub sample | global .NET chart manipulation | answer engine optimization Aspose.Cells | coding assistant chart rotation code
// Common Searches: Aspose.Cells set chart elevation C# | How to change perspective of 3D chart in Aspose.Cells | C# Aspose.Cells rotation angle for 3D chart | Save workbook after modifying 3D chart Aspose | GitHub Aspose.Cells 3D chart example
// Developer Intent: Apply specific X, Y, Z rotation values to a 3‑D chart using Aspose.Cells and persist the workbook.
// Use Cases: Generate Excel reports with pre‑oriented 3‑D charts for executive presentations. | Automate dashboard creation where chart perspective improves data readability. | Produce reusable workbook templates that include custom‑rotated 3‑D visualizations.
// AI Prompts: Write C# code with Aspose.Cells to set Elevation = 30, Perspective = 60, RotationAngle = 15 and save as MyChart.xlsx. | Explain the visual impact of Elevation, Perspective, and RotationAngle on a 3‑D chart in Aspose.Cells. | Provide step‑by‑step instructions to modify the rotation of an existing 3‑D chart in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // A ready‑to‑run C# snippet that creates a workbook with Aspose.Cells, adds a 3‑D column chart, fills it with sample data, applies Elevation = 20°, Perspective = 45, RotationAngle = 10°, and saves the result as Chart3DRotationDemo.xlsx. Ideal for .NET developers needing precise chart orientation.
    public class Chart3DRotationDemo
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
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set three‑dimensional rotation angles
                chart.Elevation = 20;          // X‑axis rotation (elevation) in degrees
                chart.Perspective = 45;        // Y‑axis rotation (perspective) (0‑100)
                chart.RotationAngle = 10;      // Z‑axis rotation around plot area

                // Save the workbook with the configured chart
                string outputPath = "Chart3DRotationDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Chart created with X=20°, Y=45°, Z=10° rotation angles. Saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
