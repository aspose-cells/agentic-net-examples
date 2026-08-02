// Title: Copy 3D Formatting Between Shapes Across Worksheets with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a rectangle shape with full ThreeDFormat settings on the first sheet, add a second rectangle on another sheet, and copy every 3D attribute—including extrusion, bevels, rotation, lighting, and camera options—from the source shape to the destination shape before saving the file.
// Keywords: Aspose.Cells | C# copy shape 3D format | ThreeDFormat transfer | shape extrusion copy | shape bevel copy | shape rotation Aspose | copy lighting settings | Aspose.Cells shape cloning | 3D shape formatting .NET
// Common Searches: How to copy ThreeDFormat from one shape to another in Aspose.Cells | Aspose.Cells copy shape 3D formatting across worksheets | Transfer extrusion, bevel and lighting properties between shapes .NET | Duplicate 3D visual style of shapes using Aspose.Cells | Copy shape 3D effects programmatically
// Developer Intent: Copy all ThreeDFormat properties from a source shape to a target shape on a different worksheet.
// Use Cases: Apply a pre‑designed 3D styled shape to many diagram elements in a multi‑sheet report. | Maintain consistent 3D appearance when cloning or moving shapes between worksheets. | Automate bulk duplication of shape visual effects without manual property setting.
// AI Prompts: Write a C# helper method that copies every ThreeDFormat property from one Aspose.Cells shape to another, including null checks. | Show C# code to copy 3D formatting between different shape types (e.g., rectangle to oval) using Aspose.Cells, handling incompatible properties. | Explain how to create a utility class for bulk copying of ThreeDFormat across a collection of shapes in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShape3DCopyDemo
{
    // Demonstrates how to create a workbook, add a rectangle shape with full ThreeDFormat settings on the first sheet, add a second rectangle on another sheet, and copy every 3D attribute—including extrusion, bevels, rotation, lighting, and camera options—from the source shape to the destination shape before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Source worksheet and shape with 3D formatting
                // -------------------------------------------------
                Worksheet srcSheet = workbook.Worksheets[0];
                // Add a rectangle shape to the source sheet
                Shape srcShape = srcSheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 100);
                // Configure 3D format properties on the source shape
                ThreeDFormat src3D = srcShape.ThreeDFormat;
                src3D.ExtrusionColor.Color = Color.Blue;
                src3D.ExtrusionHeight = 30;
                src3D.LightAngle = 45;
                src3D.Material = PresetMaterialType.Metal;
                src3D.ContourColor.Color = Color.Red;
                src3D.ContourWidth = 2;
                src3D.BottomBevelHeight = 5;
                src3D.BottomBevelWidth = 5;
                src3D.BottomBevelType = BevelType.ArtDeco;
                src3D.TopBevelHeight = 8;
                src3D.TopBevelWidth = 8;
                src3D.TopBevelType = BevelType.SoftRound;
                src3D.RotationX = 20;
                src3D.RotationY = 30;
                src3D.RotationZ = 10;
                src3D.Perspective = 15;
                src3D.PresetCameraType = PresetCameraType.PerspectiveFront;
                src3D.Lighting = LightRigType.BrightRoom;
                // Correct enum name for lighting direction
                src3D.LightingDirection = LightRigDirectionType.Top;
                src3D.Z = 5;

                // -------------------------------------------------
                // Destination worksheet and shape
                // -------------------------------------------------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");
                // Add a rectangle shape to the destination sheet (same size/position for simplicity)
                Shape destShape = destSheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 100);
                ThreeDFormat dest3D = destShape.ThreeDFormat;

                // -------------------------------------------------
                // Copy 3D format settings from source shape to destination shape
                // -------------------------------------------------
                dest3D.ExtrusionColor.Color = src3D.ExtrusionColor.Color;
                dest3D.ExtrusionHeight = src3D.ExtrusionHeight;
                dest3D.LightAngle = src3D.LightAngle;
                dest3D.Material = src3D.Material;
                dest3D.ContourColor.Color = src3D.ContourColor.Color;
                dest3D.ContourWidth = src3D.ContourWidth;
                dest3D.BottomBevelHeight = src3D.BottomBevelHeight;
                dest3D.BottomBevelWidth = src3D.BottomBevelWidth;
                dest3D.BottomBevelType = src3D.BottomBevelType;
                dest3D.TopBevelHeight = src3D.TopBevelHeight;
                dest3D.TopBevelWidth = src3D.TopBevelWidth;
                dest3D.TopBevelType = src3D.TopBevelType;
                dest3D.RotationX = src3D.RotationX;
                dest3D.RotationY = src3D.RotationY;
                dest3D.RotationZ = src3D.RotationZ;
                dest3D.Perspective = src3D.Perspective;
                dest3D.PresetCameraType = src3D.PresetCameraType;
                dest3D.Lighting = src3D.Lighting;
                dest3D.LightingDirection = src3D.LightingDirection;
                dest3D.Z = src3D.Z;

                // -------------------------------------------------
                // Save the workbook with both shapes
                // -------------------------------------------------
                workbook.Save("Shape3DCopyDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
