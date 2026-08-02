// Title: Copy Shape Reflection Settings in Aspose.Cells for .NET
// Description: Demonstrates how to create two rectangle shapes in a worksheet, apply a full‑reflection effect to the first shape, and transfer all reflection properties (type, blur, distance, direction, fade direction, size, transparency, rotation) to the second shape before saving the workbook.
// Keywords: Aspose.Cells | C# | shape reflection | copy shape properties | duplicate visual effect | Excel automation | reflection effect API | .NET workbook styling | programmatic shape formatting | Aspose.Cells tutorial
// Common Searches: Aspose.Cells copy reflection from one shape to another | C# copy shape visual effects Aspose.Cells | how to duplicate shape reflection settings in Excel using Aspose | transfer shape formatting Aspose.Cells .NET | clone shape appearance Aspose.Cells workbook
// Developer Intent: Transfer the reflection effect configuration of a source shape to another shape within the same worksheet using Aspose.Cells for .NET.
// Use Cases: Apply a consistent reflection style to multiple callout or legend shapes in a generated report. | Create a master shape template with desired visual effects and reuse it across dynamically added shapes. | Update a central shape's reflection settings and propagate the changes to all linked shapes automatically.
// AI Prompts: Write a reusable C# method that copies every reflection property from any source Aspose.Cells shape to a target shape. | Show how to extend the copy routine to include shadow, glow, and soft edge effects between shapes. | Provide code that validates the destination shape's reflection matches the source after copying.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create two rectangle shapes in a worksheet, apply a full‑reflection effect to the first shape, and transfer all reflection properties (type, blur, distance, direction, fade direction, size, transparency, rotation) to the second shape before saving the workbook.
    public class ShapeReflectionCopyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a source rectangle shape (the shape whose reflection settings will be copied)
                Shape sourceShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

                // Configure reflection settings on the source shape
                ReflectionEffect srcReflection = sourceShape.Reflection;
                srcReflection.Type = ReflectionEffectType.FullReflection4PtOffset;
                srcReflection.Blur = 5;
                srcReflection.Distance = 10;
                srcReflection.Direction = 90;
                srcReflection.FadeDirection = 45;
                srcReflection.Size = 60;
                srcReflection.Transparency = 0.3;
                srcReflection.RotWithShape = true;

                // Add a destination rectangle shape (the shape that will receive the copied settings)
                Shape destShape = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 100, 150);

                // Copy reflection settings from sourceShape to destShape
                ReflectionEffect destReflection = destShape.Reflection;
                destReflection.Type = srcReflection.Type;
                destReflection.Blur = srcReflection.Blur;
                destReflection.Distance = srcReflection.Distance;
                destReflection.Direction = srcReflection.Direction;
                destReflection.FadeDirection = srcReflection.FadeDirection;
                destReflection.Size = srcReflection.Size;
                destReflection.Transparency = srcReflection.Transparency;
                destReflection.RotWithShape = srcReflection.RotWithShape;

                // Save the workbook to verify the result
                workbook.Save("ShapeReflectionCopyDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeReflectionCopyDemo.Run();
        }
    }
}
