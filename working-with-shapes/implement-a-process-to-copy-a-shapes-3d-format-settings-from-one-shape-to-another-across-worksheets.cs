// Title: Copy 3‑D formatting from one shape to another across worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy shape ThreeDFormat from one worksheet to another in C# | How to duplicate shape rotation, depth and perspective on a different sheet with Aspose.Cells | C# example for cloning shape bevel and light rig properties using Aspose.Cells | Transfer ThreeDFormat properties between shapes via reflection in Aspose.Cells | Copy shape 3D formatting across worksheets without losing bevel settings in .NET
// Tags: shape threeDformat cloning Aspose.Cells | bevel property duplication C# | light rig transfer Aspose.Cells | reflection based shape formatting copy | cross‑worksheet shape 3D settings

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, retrieves a shape named "SourceShape" from Sheet1, creates or accesses a shape named "TargetShape" on Sheet2, and uses reflection to copy the ThreeDFormat properties—including Depth, RotationX/Y/Z, Perspective, LightRig, BevelTop, and BevelBottom—from the source shape to the target shape before saving the workbook as output.xlsx.
class Shape3DCopy
{
    static void Main()
    {
        try
        {
            const string sourceFile = "source.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the source workbook exists.
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"File not found: {sourceFile}");
                return;
            }

            // Load the workbook containing the source and target worksheets.
            Workbook workbook = new Workbook(sourceFile);

            // Access the source worksheet and the shape whose 3‑D format will be copied.
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            Shape sourceShape = sourceSheet.Shapes["SourceShape"];
            if (sourceShape == null)
            {
                Console.WriteLine("Source shape not found.");
                return;
            }

            // Access the target worksheet where the shape will receive the 3‑D format.
            Worksheet targetSheet = workbook.Worksheets["Sheet2"];
            Shape targetShape = targetSheet.Shapes["TargetShape"];
            if (targetShape == null)
            {
                // Create a new shape with the same position and size as the source shape.
                // Use a generic rectangle; the visual type can be adjusted later if needed.
                targetShape = targetSheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    sourceShape.UpperLeftRow,
                    sourceShape.UpperLeftColumn,
                    sourceShape.Top,
                    sourceShape.Left,
                    sourceShape.Height,
                    sourceShape.Width);

                if (targetShape == null)
                {
                    Console.WriteLine("Failed to create target shape.");
                    return;
                }

                targetShape.Name = "TargetShape";
            }

            // Attempt to copy 3‑D formatting using reflection (available in newer versions only).
            try
            {
                CopyThreeDFormat(sourceShape, targetShape);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ThreeDFormat copy skipped: " + ex.Message);
            }

            // Save the workbook with the updated shape.
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
        catch (Exception e)
        {
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }

    // Copies 3‑D format properties via reflection to avoid compile‑time dependencies on newer APIs.
    private static void CopyThreeDFormat(Shape srcShape, Shape tgtShape)
    {
        PropertyInfo src3DProp = srcShape.GetType().GetProperty("ThreeDFormat");
        PropertyInfo tgt3DProp = tgtShape.GetType().GetProperty("ThreeDFormat");

        if (src3DProp == null || tgt3DProp == null) return;

        object src3D = src3DProp.GetValue(srcShape);
        object tgt3D = tgt3DProp.GetValue(tgtShape);
        if (src3D == null || tgt3D == null) return;

        // Basic 3‑D properties.
        CopyProperty(src3D, tgt3D, "Depth");
        CopyProperty(src3D, tgt3D, "RotationX");
        CopyProperty(src3D, tgt3D, "RotationY");
        CopyProperty(src3D, tgt3D, "RotationZ");
        CopyProperty(src3D, tgt3D, "Perspective");

        // LightRig.
        CopyNestedProperties(src3D, tgt3D, "LightRig", new[] { "Type", "Direction", "X", "Y", "Z" });

        // BevelTop.
        CopyNestedProperties(src3D, tgt3D, "BevelTop", new[] { "Type", "Width", "Height" });

        // BevelBottom.
        CopyNestedProperties(src3D, tgt3D, "BevelBottom", new[] { "Type", "Width", "Height" });
    }

    private static void CopyProperty(object src, object tgt, string propName)
    {
        PropertyInfo srcProp = src.GetType().GetProperty(propName);
        PropertyInfo tgtProp = tgt.GetType().GetProperty(propName);
        if (srcProp != null && tgtProp != null && srcProp.CanRead && tgtProp.CanWrite)
        {
            object value = srcProp.GetValue(src);
            tgtProp.SetValue(tgt, value);
        }
    }

    private static void CopyNestedProperties(object srcParent, object tgtParent, string nestedName, string[] subProps)
    {
        PropertyInfo srcNestedProp = srcParent.GetType().GetProperty(nestedName);
        PropertyInfo tgtNestedProp = tgtParent.GetType().GetProperty(nestedName);
        if (srcNestedProp == null || tgtNestedProp == null) return;

        object srcNested = srcNestedProp.GetValue(srcParent);
        object tgtNested = tgtNestedProp.GetValue(tgtParent);
        if (srcNested == null || tgtNested == null) return;

        foreach (string subProp in subProps)
        {
            CopyProperty(srcNested, tgtNested, subProp);
        }
    }
}
