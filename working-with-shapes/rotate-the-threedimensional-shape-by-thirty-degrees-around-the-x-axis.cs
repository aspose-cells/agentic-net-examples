// Title: Rotate a 3D Shape 30° Around the X‑Axis with Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle shape, sets ThreeDFormat.RotationX to 30°, applies extrusion height and an orthographic front camera for visual effect, and saves the file as RotateShapeXDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells rotate shape X axis | ThreeDFormat RotationX .NET | 3D shape rotation Aspose.Cells | extrusion height Aspose.Cells | PresetCameraType OrthographicFront | C# Aspose.Cells 3D formatting | Excel shape 3D rotation .NET
// Common Searches: Aspose.Cells rotate shape on X axis | How to set RotationX property in Aspose.Cells | C# example for 3D shape rotation in Excel | Add extrusion height to shape Aspose.Cells | Set camera type for 3D shape Aspose.Cells
// Developer Intent: Apply a 30‑degree X‑axis rotation to a 3D rectangle in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate reports with tilted 3D graphics for better visual emphasis. | Create engineering diagrams where X‑axis tilt conveys perspective. | Produce marketing dashboards that showcase 3D objects with realistic depth. | Combine RotationX, extrusion height, and orthographic camera to simulate 3D models in Excel.
// AI Prompts: Show how to rotate a shape around the Y axis with Aspose.Cells for .NET. | Provide code to animate a 3D shape rotating incrementally on the X axis across multiple worksheets. | Explain how to reset ThreeDFormat properties to defaults after applying a rotation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape, sets ThreeDFormat.RotationX to 30°, applies extrusion height and an orthographic front camera for visual effect, and saves the file as RotateShapeXDemo.xlsx using Aspose.Cells for .NET.
    public class RotateShapeXDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: type, upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 2, 200, 100, 100);

                // Access the shape's 3D formatting object
                ThreeDFormat threeDFormat = shape.ThreeDFormat;

                // Rotate the shape 30 degrees around the X‑axis
                threeDFormat.RotationX = 30;

                // Set additional 3D properties so the rotation is visible
                threeDFormat.ExtrusionHeight = 20;
                threeDFormat.PresetCameraType = PresetCameraType.OrthographicFront;

                // Save the workbook with the rotated shape
                workbook.Save("RotateShapeXDemo.xlsx");
                Console.WriteLine("Workbook saved as RotateShapeXDemo.xlsx");
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
            RotateShapeXDemo.Run();
        }
    }
}
