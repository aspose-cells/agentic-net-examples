// Title: Copy 3D formatting from a source shape to a target shape on another worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy ThreeDFormat from one shape to another worksheet | how to transfer shape 3D properties between worksheets using Aspose.Cells | C# example for cloning shape 3D formatting in Excel with Aspose.Cells | copy shape threeDformat across worksheets Aspose.Cells .NET | duplicate shape 3D format settings in a different sheet using Aspose.Cells
// Tags: copy shape threeDformat Aspose.Cells | transfer shape 3d properties Excel .NET | clone shape 3d formatting across worksheets | Aspose.Cells shape 3d format reflection | copy shape formatting between worksheets C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Shape3DCopyDemo
{
    // Loads an Excel workbook, finds a source shape and a target shape (creating the target worksheet and a placeholder rectangle if needed), then uses reflection to copy all writable ThreeDFormat properties from the source shape to the target shape, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Expected arguments:
            // 0 - input workbook path
            // 1 - source worksheet name
            // 2 - source shape name
            // 3 - target worksheet name
            // 4 - target shape name (if shape does not exist it will be created)
            // 5 - output workbook path
            if (args.Length != 6)
            {
                Console.WriteLine("Usage: Shape3DCopyDemo <input.xlsx> <SourceSheet> <SourceShapeName> <TargetSheet> <TargetShapeName> <output.xlsx>");
                return;
            }

            string inputPath = args[0];
            string sourceSheetName = args[1];
            string sourceShapeName = args[2];
            string targetSheetName = args[3];
            string targetShapeName = args[4];
            string outputPath = args[5];

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);

                // Get source worksheet
                Worksheet srcSheet = workbook.Worksheets[sourceSheetName];
                if (srcSheet == null)
                {
                    Console.WriteLine($"Source worksheet \"{sourceSheetName}\" not found.");
                    return;
                }

                // Get source shape
                Shape srcShape = GetShapeByName(srcSheet, sourceShapeName);
                if (srcShape == null)
                {
                    Console.WriteLine($"Source shape \"{sourceShapeName}\" not found in worksheet \"{sourceSheetName}\".");
                    return;
                }

                // Get or create target worksheet
                Worksheet tgtSheet = workbook.Worksheets[targetSheetName];
                if (tgtSheet == null)
                {
                    Console.WriteLine($"Target worksheet \"{targetSheetName}\" not found. Creating new worksheet.");
                    tgtSheet = workbook.Worksheets.Add(targetSheetName);
                }

                // Get or create target shape
                Shape tgtShape = GetShapeByName(tgtSheet, targetShapeName);
                if (tgtShape == null)
                {
                    Console.WriteLine($"Target shape \"{targetShapeName}\" not found. Adding a rectangle as placeholder.");
                    // Parameters: upper left row, upper left column, upper left offsetX, upper left offsetY, height, width
                    RectangleShape rect = tgtSheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);
                    rect.Name = targetShapeName;
                    tgtShape = rect;
                }

                // Ensure both shapes support 3D format
                if (srcShape.ThreeDFormat == null || tgtShape.ThreeDFormat == null)
                {
                    Console.WriteLine("One of the shapes does not support 3D formatting.");
                    return;
                }

                // Copy 3D format settings
                CopyThreeDFormat(srcShape.ThreeDFormat, tgtShape.ThreeDFormat);

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"3D format copied successfully. Output saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static Shape GetShapeByName(Worksheet sheet, string shapeName)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                if (string.Equals(shape.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }
            }
            return null;
        }

        private static void CopyThreeDFormat(ThreeDFormat source, ThreeDFormat destination)
        {
            if (source == null || destination == null) return;

            PropertyInfo[] properties = typeof(ThreeDFormat).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in properties)
            {
                if (!prop.CanWrite) continue;

                try
                {
                    object value = prop.GetValue(source);
                    // Attempt shallow clone for complex reference types
                    if (value != null && !prop.PropertyType.IsValueType && prop.PropertyType != typeof(string))
                    {
                        MethodInfo cloneMethod = prop.PropertyType.GetMethod("Clone", BindingFlags.Public | BindingFlags.Instance);
                        if (cloneMethod != null)
                        {
                            value = cloneMethod.Invoke(value, null);
                        }
                    }
                    prop.SetValue(destination, value);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not copy property \"{prop.Name}\" – {ex.Message}");
                }
            }
        }
    }
}
