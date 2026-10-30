// Title: Copy reflection effect settings from one shape to another in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to duplicate shape reflection effect in Aspose.Cells C# | copy shape EffectFormat reflection properties between shapes in Excel using Aspose.Cells | Aspose.Cells C# copy reflection settings from source shape to target shape | use .NET reflection to transfer shape visual effects in an Excel file with Aspose.Cells | clone shape reflection attributes (type, transparency, size) in a workbook via Aspose.Cells
// Tags: shape reflection copy Aspose.Cells | EffectFormat reflection transfer C# | copy shape visual effects Excel .NET | Aspose.Cells shape effect duplication | reflection property cloning workbook

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsReflectionCopy
{
    // The example loads an Excel workbook, locates two shapes by name, and uses .NET reflection to access each shape's EffectFormat and Reflection objects. It then copies reflection properties such as Type, Transparency, Size, Blur, Distance, Angle, Alignment, OffsetX, and OffsetY from the source shape to the target shape before saving the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";
            const string sourceShapeName = "SourceShapeName";
            const string targetShapeName = "TargetShapeName";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var sheet = workbook.Worksheets[0];

                // Retrieve source and target shapes by name
                var sourceShape = sheet.Shapes[sourceShapeName];
                var targetShape = sheet.Shapes[targetShapeName];

                if (sourceShape == null)
                {
                    Console.WriteLine($"Source shape \"{sourceShapeName}\" not found.");
                    return;
                }

                if (targetShape == null)
                {
                    Console.WriteLine($"Target shape \"{targetShapeName}\" not found.");
                    return;
                }

                // Copy reflection properties using reflection (to support versions without direct EffectFormat API)
                CopyReflectionProperties(sourceShape, targetShape);

                // Ensure output directory exists
                var outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void CopyReflectionProperties(Shape source, Shape target)
        {
            try
            {
                // Get EffectFormat property info
                PropertyInfo sourceEffectProp = source.GetType().GetProperty("EffectFormat");
                PropertyInfo targetEffectProp = target.GetType().GetProperty("EffectFormat");

                if (sourceEffectProp == null || targetEffectProp == null)
                {
                    Console.WriteLine("EffectFormat property not available in this Aspose.Cells version.");
                    return;
                }

                // Get EffectFormat instances
                object sourceEffect = sourceEffectProp.GetValue(source);
                object targetEffect = targetEffectProp.GetValue(target);

                if (sourceEffect == null || targetEffect == null)
                {
                    Console.WriteLine("EffectFormat instance is null.");
                    return;
                }

                // Get Reflection property info
                PropertyInfo sourceReflectionProp = sourceEffect.GetType().GetProperty("Reflection");
                PropertyInfo targetReflectionProp = targetEffect.GetType().GetProperty("Reflection");

                if (sourceReflectionProp == null || targetReflectionProp == null)
                {
                    Console.WriteLine("Reflection property not available in this Aspose.Cells version.");
                    return;
                }

                // Get Reflection instances
                object sourceReflection = sourceReflectionProp.GetValue(sourceEffect);
                object targetReflection = targetReflectionProp.GetValue(targetEffect);

                if (sourceReflection == null || targetReflection == null)
                {
                    Console.WriteLine("Reflection instance is null.");
                    return;
                }

                // List of reflection property names to copy
                string[] reflectionProps = new[]
                {
                    "Type", "Transparency", "Size", "Blur", "Distance",
                    "Angle", "Alignment", "OffsetX", "OffsetY"
                };

                // Copy each property value from source to target
                foreach (string propName in reflectionProps)
                {
                    PropertyInfo srcProp = sourceReflection.GetType().GetProperty(propName);
                    PropertyInfo tgtProp = targetReflection.GetType().GetProperty(propName);

                    if (srcProp != null && tgtProp != null && tgtProp.CanWrite)
                    {
                        object value = srcProp.GetValue(sourceReflection);
                        tgtProp.SetValue(targetReflection, value);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to copy reflection properties: {ex.Message}");
            }
        }
    }
}
