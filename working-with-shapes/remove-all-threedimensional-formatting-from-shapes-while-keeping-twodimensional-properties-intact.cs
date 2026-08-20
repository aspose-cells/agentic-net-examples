// Title: Remove 3D Formatting from Shapes and Chart Series in Aspose.Cells for .NET
// Description: Shows how to clear all three‑dimensional effects from worksheet shapes and chart series while preserving their two‑dimensional appearance using Aspose.Cells for .NET. The sample creates a workbook, adds a rectangle with extrusion and a 3D column chart, then uses ClearFormat3D and resets ThreeDFormat properties before saving the file.
// Keywords: Aspose.Cells 3D shape removal | ClearFormat3D C# | reset ThreeDFormat Aspose.Cells | remove extrusion Aspose.Cells | disable 3D chart series Aspose.Cells | C# Excel shape formatting | Aspose.Cells workbook cleanup | remove bevel Aspose.Cells
// Common Searches: How to clear 3D formatting from shapes in Aspose.Cells .NET | Remove extrusion from Excel shapes using Aspose.Cells | Clear 3D effects from chart series Aspose.Cells C# | Reset ThreeDFormat properties to default in Aspose.Cells | Strip 3D formatting from an Excel workbook programmatically
// Developer Intent: Eliminate all three‑dimensional formatting from worksheet shapes and chart series while keeping their flat, two‑dimensional visual properties unchanged.
// Use Cases: Prepare Excel files for printing by removing 3D shadows, bevels, and extrusion. | Standardize exported workbooks by resetting shape lighting and material settings. | Clean legacy spreadsheets that contain unwanted 3D effects before distribution.
// AI Prompts: Write C# code with Aspose.Cells that iterates through every worksheet and sets Shape.ThreeDFormat.ExtrusionHeight to zero for each shape. | Provide an example that calls ShapePropertyCollection.ClearFormat3D on all chart series and then saves the workbook. | Explain how to reset all ThreeDFormat attributes (bevel, material, lighting, rotation) to their defaults for shapes in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to clear all three‑dimensional effects from worksheet shapes and chart series while preserving their two‑dimensional appearance using Aspose.Cells for .NET. The sample creates a workbook, adds a rectangle with extrusion and a 3D column chart, then uses ClearFormat3D and resets ThreeDFormat properties before saving the file.
    public class Remove3DFormattingFromShapes
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Sample shapes with 3‑D formatting (for demonstration)
            // -------------------------------------------------
            // Add a regular shape and give it some 3‑D extrusion
            Shape rectangle = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 100);
            rectangle.ThreeDFormat.ExtrusionHeight = 30;          // 3‑D property
            rectangle.ThreeDFormat.TopBevelHeight = 5;           // 3‑D property

            // Add a 3‑D chart and configure its series 3‑D format
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 15, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Populate chart data
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            chart.NSeries.Add("A1:A2", true);

            // Access the first series shape properties and set 3‑D formatting
            Series series = chart.NSeries[0];
            ShapePropertyCollection seriesShapeProps = series.ShapeProperties;
            seriesShapeProps.Format3D.TopBevel.Height = 8;
            // Note: ExtrusionHeight is not available for series Format3D; omitted.

            // -------------------------------------------------
            // Remove all 3‑D formatting while preserving 2‑D properties
            // -------------------------------------------------
            // 1. Clear 3‑D formatting from all chart series shapes
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    foreach (Series s in ch.NSeries)
                    {
                        ShapePropertyCollection spc = s.ShapeProperties;
                        spc.ClearFormat3D();
                    }
                }
            }

            // 2. Reset 3‑D properties of regular shapes (no dedicated Clear method)
            //    Setting the key 3‑D attributes to their default values effectively removes the effect.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Shape shp in ws.Shapes)
                {
                    if (shp.ThreeDFormat != null)
                    {
                        shp.ThreeDFormat.ExtrusionHeight = 0;
                        shp.ThreeDFormat.ExtrusionColor.Color = Color.Empty;
                        shp.ThreeDFormat.TopBevelHeight = 0;
                        shp.ThreeDFormat.TopBevelWidth = 0;
                        shp.ThreeDFormat.BottomBevelHeight = 0;
                        shp.ThreeDFormat.BottomBevelWidth = 0;
                        shp.ThreeDFormat.Material = PresetMaterialType.Plastic;
                        shp.ThreeDFormat.LightAngle = 0;
                        shp.ThreeDFormat.RotationX = 0;
                        shp.ThreeDFormat.RotationY = 0;
                        shp.ThreeDFormat.RotationZ = 0;
                    }
                }
            }

            // -------------------------------------------------
            // Save the workbook with 3‑D formatting removed
            // -------------------------------------------------
            string outputPath = "Removed3DFormatting.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
