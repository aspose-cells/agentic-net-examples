// Title: Rotate a 3‑D Shape 30° Around the X‑Axis with Aspose.Cells for .NET (C# Example)
// Description: This C# sample creates a workbook, adds a rectangle shape, accesses its ThreeDFormat, sets RotationX to 30°, applies extrusion height, contour width, and an orthographic front camera, then saves the file as an Excel workbook with the rotated 3‑D shape.
// Keywords: Aspose.Cells rotate shape X axis | ThreeDFormat RotationX C# | Excel 3D shape transformation | C# Aspose.Cells extrusion height | preset camera Aspose.Cells shape | .NET Excel shape rotation example | 3‑D rectangle shape Aspose.Cells | shape rotation X‑axis Aspose.Cells
// Common Searches: how to rotate a shape around the X axis using Aspose.Cells for .NET | Aspose.Cells C# example for ThreeDFormat RotationX | set extrusion height and contour width for Excel shapes in C# | apply orthographic front camera to a 3‑D shape with Aspose.Cells | rotate 3‑D rectangle shape in Excel programmatically
// Developer Intent: Apply a 30° X‑axis rotation to a 3‑D rectangle shape and save the workbook using Aspose.Cells for .NET.
// Use Cases: Create isometric diagrams in Excel by tilting shapes for clearer data visualization. | Enhance financial or engineering reports with 3‑D shapes that emphasize key metrics. | Prepare presentation‑ready Excel files where shapes need custom extrusion and rotation.
// AI Prompts: Show me how to rotate a shape around the Y axis with Aspose.Cells for .NET. | Explain how to combine RotationX with lighting and material settings for a 3‑D shape in Aspose.Cells. | Provide a C# code snippet that animates a shape rotating on multiple axes in an Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, adds a rectangle shape, accesses its ThreeDFormat, sets RotationX to 30°, applies extrusion height, contour width, and an orthographic front camera, then saves the file as an Excel workbook with the rotated 3‑D shape.
    public class RotateShapeAroundXAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: drawing type, upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 2, 200, 100, 100);

                // Access the 3‑D format of the shape
                ThreeDFormat threeDFormat = shape.ThreeDFormat;

                // Rotate the shape 30 degrees around the X‑axis
                threeDFormat.RotationX = 30;

                // Set additional 3‑D properties so the rotation is visible
                threeDFormat.ExtrusionHeight = 20;          // give the shape depth
                threeDFormat.ContourWidth = 2;              // outline thickness
                threeDFormat.PresetCameraType = PresetCameraType.OrthographicFront;

                // Save the workbook with the rotated shape
                string outputPath = "ShapeRotated30DegAroundX.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Shape rotated 30° around X‑axis and saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RotateShapeAroundXAxis.Run();
        }
    }
}
