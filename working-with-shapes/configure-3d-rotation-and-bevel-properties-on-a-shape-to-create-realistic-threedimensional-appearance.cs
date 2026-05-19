using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCells3DRotationAndBevel
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate 3D formatting
            // Parameters: drawing type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 2, 2, 150, 200);
            shape.Text = "3D Shape";

            // Access the ThreeDFormat of the shape
            ThreeDFormat threeD = shape.ThreeDFormat;

            // Set extrusion (depth) to give the shape thickness
            threeD.ExtrusionHeight = 30;          // points

            // Configure realistic rotation around X, Y and Z axes
            threeD.RotationX = 30;                // tilt forward/backward
            threeD.RotationY = 40;                // tilt left/right
            threeD.RotationZ = 15;                // spin around its center

            // Adjust perspective to control the viewing angle (0‑120 degrees)
            threeD.Perspective = 45;

            // Set material and lighting to enhance 3‑D appearance
            threeD.Material = PresetMaterialType.Metal;
            threeD.LightAngle = 45;               // direction of the light source
            threeD.Lighting = LightRigType.ThreePoint;
            threeD.LightingDirection = LightRigDirectionType.Top;

            // Configure top bevel (makes the front edge look rounded)
            threeD.TopBevelType = BevelType.SoftRound;
            threeD.TopBevelWidth = 10;
            threeD.TopBevelHeight = 10;

            // Configure bottom bevel (adds depth to the back edge)
            threeD.BottomBevelType = BevelType.Divot;
            threeD.BottomBevelWidth = 5;
            threeD.BottomBevelHeight = 5;

            // Optional: add contour to emphasize edges
            threeD.ContourWidth = 1.5;
            threeD.ContourColor.Color = Color.DarkGray;

            // Save the workbook (lifecycle save)
            workbook.Save("ThreeDRotationBevelDemo.xlsx");
        }
    }
}