// Title: Copy Shape 3D Formatting Between Worksheets with Aspose.Cells for .NET
// Description: Demonstrates how to duplicate every ThreeDFormat property (extrusion, bevels, contour, lighting, material, rotation, Z‑distance) from a source shape to a target shape on another worksheet, then save the workbook as an Excel file.
// Keywords: Aspose.Cells copy shape 3D format | ThreeDFormat transfer .NET | duplicate shape extrusion Aspose | copy bevel rotation Excel shape | shape lighting material Aspose.Cells
// Common Searches: copy 3d format between shapes Aspose.Cells | transfer shape extrusion settings across worksheets | duplicate ThreeDFormat properties in .NET | Aspose.Cells copy bevel and lighting | how to clone shape 3d appearance in Excel
// Developer Intent: Programmatically clone all 3D formatting attributes from one Excel shape to another on a different worksheet.
// Use Cases: Apply a consistent 3D visual theme to shapes across multiple report sheets. | Reuse a styled rectangle or callout in template‑driven workbooks without manual reformatting. | Synchronize shape appearance when generating dashboards that share the same 3D effects.
// AI Prompts: Create a generic method that copies ThreeDFormat properties between any two Aspose.Cells Shape objects, handling nulls and optional filters. | Show code to loop through a template worksheet’s shapes and replicate their 3D formats to several destination worksheets. | Provide robust error handling and logging for the CopyThreeDFormat routine when shapes are missing or the workbook is read‑only.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCopy3DFormat
{
    // Demonstrates how to duplicate every ThreeDFormat property (extrusion, bevels, contour, lighting, material, rotation, Z‑distance) from a source shape to a target shape on another worksheet, then save the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the source and destination worksheets
            Workbook workbook = new Workbook(); // create workbook
            Worksheet sourceSheet = workbook.Worksheets[0];
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Add a rectangle shape to the source worksheet and configure its 3D format
            Shape sourceShape = sourceSheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 100);
            ThreeDFormat srcFmt = sourceShape.ThreeDFormat;
            srcFmt.ExtrusionColor.Color = Color.Blue;
            srcFmt.LightAngle = 45;
            srcFmt.ContourWidth = 2;
            srcFmt.ContourColor.Color = Color.Red;
            srcFmt.ExtrusionHeight = 30;
            srcFmt.Material = PresetMaterialType.Metal;
            srcFmt.RotationX = 15;
            srcFmt.RotationY = 30;
            srcFmt.RotationZ = 45;

            // Add a rectangle shape to the destination worksheet (same size/position)
            Shape destShape = destSheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 100);

            // Copy all 3D format settings from sourceShape to destShape
            CopyThreeDFormat(sourceShape, destShape);

            // Save the workbook
            workbook.Save("Copy3DFormatDemo.xlsx");
        }

        // Copies every 3D property from one shape to another
        static void CopyThreeDFormat(Shape source, Shape target)
        {
            ThreeDFormat src = source.ThreeDFormat;
            ThreeDFormat dst = target.ThreeDFormat;

            // Bevels
            dst.BottomBevelHeight = src.BottomBevelHeight;
            dst.BottomBevelWidth = src.BottomBevelWidth;
            dst.BottomBevelType = src.BottomBevelType;

            dst.TopBevelHeight = src.TopBevelHeight;
            dst.TopBevelWidth = src.TopBevelWidth;
            dst.TopBevelType = src.TopBevelType;

            // Contour
            dst.ContourWidth = src.ContourWidth;
            dst.ContourColor.Color = src.ContourColor.Color;

            // Extrusion
            dst.ExtrusionHeight = src.ExtrusionHeight;
            dst.ExtrusionColor.Color = src.ExtrusionColor.Color;

            // Lighting and material
            dst.LightAngle = src.LightAngle;
            dst.Lighting = src.Lighting;
            dst.LightingDirection = src.LightingDirection;
            dst.Material = src.Material;
            dst.Perspective = src.Perspective;
            dst.PresetCameraType = src.PresetCameraType;

            // Rotation
            dst.RotationX = src.RotationX;
            dst.RotationY = src.RotationY;
            dst.RotationZ = src.RotationZ;

            // Z distance
            dst.Z = src.Z;
        }
    }
}
