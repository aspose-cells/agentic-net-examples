using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class CopyReflectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the source shape
            Shape sourceShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);
            // Get the existing reflection effect (read‑only property) and configure it
            ReflectionEffect sourceReflection = sourceShape.Reflection;
            sourceReflection.Type = ReflectionEffectType.FullReflection4PtOffset;
            sourceReflection.Blur = 5;
            sourceReflection.Distance = 10;
            sourceReflection.Direction = 90;
            sourceReflection.FadeDirection = 90;
            sourceReflection.Size = 50;
            sourceReflection.Transparency = 0.5;
            sourceReflection.RotWithShape = true;

            // Add the target shape
            Shape targetShape = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 100, 150);
            // Get the reflection effect for the target shape
            ReflectionEffect targetReflection = targetShape.Reflection;

            // Copy each reflection property from the source shape to the target shape
            targetReflection.Type = sourceReflection.Type;
            targetReflection.Blur = sourceReflection.Blur;
            targetReflection.Distance = sourceReflection.Distance;
            targetReflection.Direction = sourceReflection.Direction;
            targetReflection.FadeDirection = sourceReflection.FadeDirection;
            targetReflection.Size = sourceReflection.Size;
            targetReflection.Transparency = sourceReflection.Transparency;
            targetReflection.RotWithShape = sourceReflection.RotWithShape;

            // Save the workbook to a file
            workbook.Save("CopyReflectionDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}