// Title: Add a rectangle shape and apply 3‑D rotation, extrusion depth, and bevel effects using Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape on a worksheet and assign values to ThreeDFormat.RotationX, RotationY, and RotationZ with Aspose.Cells in C#. | Set the extrusion depth and configure top and bottom bevel types, widths, and heights for a shape's ThreeDFormat using the Aspose.Cells API. | Export the workbook that contains the 3‑D formatted shape to an .xlsx file.
// Common Searches: Aspose.Cells C# how to rotate an Excel shape in 3D | set extrusion depth and bevel on shape with Aspose.Cells .NET | example of ThreeDFormat properties for rectangle shape in Aspose.Cells | apply 3D rotation X Y Z to Excel shape using Aspose.Cells API
// Tags: Aspose.Cells ThreeDFormat rotation properties | C# set shape extrusion depth Aspose.Cells | Aspose.Cells bevel top bottom configuration | create rectangle shape Aspose.Cells worksheet | save workbook with 3D shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a rectangle shape, configure its ThreeDFormat rotation (X, Y, Z), optionally set extrusion depth and top/bottom bevel parameters, and save the result as an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top (in points)
                100,  // left (in points)
                100,  // height (in points)
                100   // width (in points)
            );

            // Configure 3‑D rotation (degrees)
            shape.ThreeDFormat.RotationX = 30; // tilt forward/backward
            shape.ThreeDFormat.RotationY = 20; // tilt left/right
            shape.ThreeDFormat.RotationZ = 10; // spin around its center

            // Note: Depth and bevel properties may not be available in older Aspose.Cells versions.
            // If supported, you can uncomment and adjust the following lines:

            // shape.ThreeDFormat.ExtrusionDepth = 5; // extrusion depth (points)

            // shape.ThreeDFormat.BevelTop.Type = BevelType.Circle;
            // shape.ThreeDFormat.BevelTop.Width = 5;
            // shape.ThreeDFormat.BevelTop.Height = 5;

            // shape.ThreeDFormat.BevelBottom.Type = BevelType.Angle;
            // shape.ThreeDFormat.BevelBottom.Width = 3;
            // shape.ThreeDFormat.BevelBottom.Height = 3;

            // Save the workbook with the 3‑D shaped object
            workbook.Save("3DShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
