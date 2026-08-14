// Title: Copy shape reflection effect between rectangles using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add two rectangle shapes on the same worksheet, configure a full reflection effect on the source shape, and transfer all ReflectionEffect properties (type, blur, distance, direction, fade direction, size, transparency, RotWithShape) to the destination shape before saving the file.
// Keywords: Aspose.Cells | C# | shape reflection | ReflectionEffect | copy shape properties | Aspose.Cells.Drawing | worksheet shapes | CopyShapeReflectionDemo | duplicate visual style | Aspose.Cells example
// Common Searches: copy reflection effect Aspose.Cells C# | transfer shape reflection properties in .NET | Aspose.Cells copy shape visual settings | how to duplicate reflection on multiple shapes Aspose.Cells | copy shape reflection between rectangles worksheet
// Developer Intent: Copy all reflection effect settings from a source shape to another shape within the same worksheet using Aspose.Cells for .NET.
// Use Cases: Apply a consistent reflection style to a series of template shapes in a generated report. | Maintain uniform visual effects when programmatically cloning chart legends or icons. | Synchronize reflection attributes after shape repositioning to preserve design consistency.
// AI Prompts: Write a C# method that receives two Aspose.Cells Shape objects and copies every ReflectionEffect property from the first to the second. | Show how to copy additional visual effects such as Glow, Shadow, and SoftEdges between shapes with Aspose.Cells. | Explain the steps required to copy reflection settings when the source and destination shapes reside on different worksheets in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add two rectangle shapes on the same worksheet, configure a full reflection effect on the source shape, and transfer all ReflectionEffect properties (type, blur, distance, direction, fade direction, size, transparency, RotWithShape) to the destination shape before saving the file.
    public class CopyShapeReflectionDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the source shape (rectangle) and configure its reflection effect
            Shape sourceShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);
            ReflectionEffect sourceReflection = sourceShape.Reflection;
            sourceReflection.Type = ReflectionEffectType.FullReflection4PtOffset;
            sourceReflection.Blur = 5;
            sourceReflection.Distance = 10;
            sourceReflection.Direction = 90;
            sourceReflection.FadeDirection = 90;
            sourceReflection.Size = 50;
            sourceReflection.Transparency = 0.5;
            sourceReflection.RotWithShape = true;

            // Add the destination shape (rectangle) where the reflection settings will be copied
            Shape destShape = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 100, 150);
            ReflectionEffect destReflection = destShape.Reflection;

            // Copy each reflection property from source to destination
            destReflection.Type = sourceReflection.Type;
            destReflection.Blur = sourceReflection.Blur;
            destReflection.Distance = sourceReflection.Distance;
            destReflection.Direction = sourceReflection.Direction;
            destReflection.FadeDirection = sourceReflection.FadeDirection;
            destReflection.Size = sourceReflection.Size;
            destReflection.Transparency = sourceReflection.Transparency;
            destReflection.RotWithShape = sourceReflection.RotWithShape;

            // Save the workbook to a file
            string outputPath = "CopyShapeReflectionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
