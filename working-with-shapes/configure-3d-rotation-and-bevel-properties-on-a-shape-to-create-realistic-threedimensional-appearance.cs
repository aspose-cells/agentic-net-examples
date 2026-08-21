// Title: C# – Apply 3D Rotation, Extrusion, Bevel, Material & Lighting to a Shape with Aspose.Cells
// Description: Demonstrates how to add a rectangle shape to an Excel workbook and use the Shape.ThreeDFormat API to set extrusion depth, rotate on X/Y/Z axes, adjust perspective, apply top and bottom bevels, choose a metal material, and configure three‑point lighting before saving the file.
// Keywords: Aspose.Cells 3D shape | C# shape rotation X Y Z | extrusion height Aspose.Cells | shape bevel Aspose.Cells | material preset metal | three‑point lighting Excel | ThreeDFormat API
// Common Searches: Aspose.Cells rotate shape on X axis | how to add bevel to Excel shape using C# | set extrusion height for rectangle in Aspose.Cells | apply material and lighting to 3D shape Aspose | C# example for 3D shape formatting in Excel
// Developer Intent: Create a realistic 3‑D visual effect for a worksheet shape by configuring rotation, depth, bevels, material, and lighting with Aspose.Cells for .NET.
// Use Cases: Design a 3‑D button with rounded top bevel and metallic finish for an interactive dashboard. | Build an inset panel with custom X/Y/Z rotation and divot bottom bevel for a report sidebar. | Enhance a chart legend with perspective and three‑point lighting to improve depth perception.
// AI Prompts: Generate C# code that adds a cylinder shape to an Aspose.Cells workbook, rotates it 45° on the X‑axis, applies a soft‑round top bevel, and uses a glass material with two‑point lighting. | Show how to change LightAngle and LightingDirection of a shape's ThreeDFormat based on a user‑selected option in a C# application.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Demonstrates how to add a rectangle shape to an Excel workbook and use the Shape.ThreeDFormat API to set extrusion depth, rotate on X/Y/Z axes, adjust perspective, apply top and bottom bevels, choose a metal material, and configure three‑point lighting before saving the file.
class ThreeDRotationBevelDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 1, 1, 200, 100);
        shape.Text = "3D Shape";

        // Access the 3D formatting object of the shape
        ThreeDFormat threeD = shape.ThreeDFormat;

        // Give the shape depth
        threeD.ExtrusionHeight = 30;

        // Rotate the shape around X, Y and Z axes
        threeD.RotationX = 30; // tilt forward
        threeD.RotationY = 20; // rotate sideways
        threeD.RotationZ = 15; // spin

        // Set the viewing perspective
        threeD.Perspective = 45;

        // Configure top bevel for a smooth rounded edge
        threeD.TopBevelType = BevelType.SoftRound;
        threeD.TopBevelWidth = 8;
        threeD.TopBevelHeight = 8;

        // Configure bottom bevel for a subtle inset
        threeD.BottomBevelType = BevelType.Divot;
        threeD.BottomBevelWidth = 4;
        threeD.BottomBevelHeight = 4;

        // Apply material and lighting to enhance realism
        threeD.Material = PresetMaterialType.Metal;
        threeD.LightAngle = 60;
        threeD.Lighting = LightRigType.ThreePoint;
        threeD.LightingDirection = LightRigDirectionType.Top;

        // Save the workbook with the 3D formatted shape
        workbook.Save("ThreeDRotationBevelDemo.xlsx");
    }
}
