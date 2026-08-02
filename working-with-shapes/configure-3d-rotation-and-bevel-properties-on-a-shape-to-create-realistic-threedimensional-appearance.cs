// Title: Apply 3D Rotation, Extrusion, and Bevel Effects to an Excel Shape with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a rectangle shape, and uses the ThreeDFormat API to set RotationX/Y/Z, Perspective, ExtrusionHeight, top and bottom bevels, metal material, and a three‑point light rig, producing a realistic 3‑D appearance before saving the file.
// Keywords: Aspose.Cells | C# shape 3D | ThreeDFormat | RotationX | RotationY | RotationZ | Perspective | ExtrusionHeight | BevelType | TopBevel | BottomBevel | PresetMaterialType | LightRig | Excel shape styling | .NET Excel 3D | Aspose.Cells example
// Common Searches: Aspose.Cells set shape rotation X Y Z | How to add bevel to Excel shape using Aspose.Cells | Extrusion depth for rectangle shape Aspose.Cells .NET | Apply material and lighting to Excel shape with Aspose.Cells | ThreeDFormat properties tutorial
// Developer Intent: Create a rectangle shape in an Excel workbook and configure its 3‑D rotation, extrusion depth, bevels, material, and lighting to achieve a lifelike three‑dimensional look.
// Use Cases: Generate reports that include visually rich 3‑D shapes for dashboards. | Design marketing templates where shapes need realistic shading and depth. | Automate the creation of instructional worksheets that highlight objects with extrusion and bevel effects.
// AI Prompts: Provide C# code that adds a rectangle to a worksheet and configures ThreeDFormat with rotation, extrusion, bevel types, metal material, and a three‑point light rig using Aspose.Cells. | Explain the impact of RotationX, RotationY, RotationZ, and Perspective on the visual orientation of an Excel shape created with Aspose.Cells. | Suggest optimal bevel and material combinations for a realistic 3‑D effect on shapes in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a rectangle shape, and uses the ThreeDFormat API to set RotationX/Y/Z, Perspective, ExtrusionHeight, top and bottom bevels, metal material, and a three‑point light rig, producing a realistic 3‑D appearance before saving the file.
    public class ShapeThreeDRotationAndBevelDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: drawing type, upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    2,   // upper left row
                    2,   // upper left column
                    2,   // top offset (points)
                    2,   // left offset (points)
                    150, // height (points)
                    200  // width (points)
                );

                // Access the 3D format of the shape
                ThreeDFormat threeD = shape.ThreeDFormat;

                // Configure realistic 3‑D rotation
                threeD.RotationX = 30;      // rotate 30° around X‑axis
                threeD.RotationY = 20;      // rotate 20° around Y‑axis
                threeD.RotationZ = 15;      // rotate 15° around Z‑axis
                threeD.Perspective = 45;    // set viewing perspective to 45°

                // Set extrusion to give depth
                threeD.ExtrusionHeight = 25; // depth of extrusion in points

                // Configure top bevel (soft round for smooth edges)
                threeD.TopBevelType = BevelType.SoftRound;
                threeD.TopBevelWidth = 10;   // width of top bevel
                threeD.TopBevelHeight = 10;  // height of top bevel

                // Configure bottom bevel (divot for subtle inset)
                threeD.BottomBevelType = BevelType.Divot;
                threeD.BottomBevelWidth = 5;
                threeD.BottomBevelHeight = 5;

                // Optional: set material and lighting to enhance realism
                threeD.Material = PresetMaterialType.Metal;
                threeD.LightAngle = 45;          // direction of light source
                threeD.Lighting = LightRigType.ThreePoint;
                threeD.LightingDirection = LightRigDirectionType.Top;

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeThreeDRotationAndBevelDemo.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeThreeDRotationAndBevelDemo.Run();
        }
    }
}
