// Title: Copy reflection effect settings from one shape to another within the same worksheet using Aspose.Cells for .NET
// Common Searches: how to copy reflection effect from one shape to another with Aspose.Cells C# | Aspose.Cells transfer shape reflection properties in the same worksheet | duplicate shape visual effects (reflection) using Aspose.Cells .NET | copy shape formatting reflection settings programmatically in Aspose.Cells | clone reflection effect between rectangle and oval shapes in a workbook
// Tags: copy shape reflection settings Aspose.Cells | transfer reflection effect between shapes .NET | shape visual effect cloning Aspose.Cells API | duplicate shape formatting reflection worksheet | Aspose.Cells reflection property copy | clone shape effects in Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to copy all reflection effect properties from a source rectangle shape to a target oval auto shape in the same worksheet using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a source rectangle shape and configure its reflection effect
            Shape sourceShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);
            ReflectionEffect sourceReflection = sourceShape.Reflection;
            sourceReflection.Type = ReflectionEffectType.HalfReflectionTouching;
            sourceReflection.Blur = 4.5;
            sourceReflection.Distance = 12;
            sourceReflection.Direction = 90;
            sourceReflection.FadeDirection = 45;
            sourceReflection.Size = 60;
            sourceReflection.Transparency = 0.3;
            sourceReflection.RotWithShape = true;

            // Add a target ellipse shape using AddAutoShape (AutoShapeType.Oval)
            Shape targetShape = worksheet.Shapes.AddAutoShape(
                AutoShapeType.Oval, // ellipse shape
                5, 0,               // upper left row, column
                5, 0,               // top, left offsets
                150, 100);          // height, width

            ReflectionEffect targetReflection = targetShape.Reflection;

            // Copy all reflection properties from source to target
            CopyReflectionSettings(sourceReflection, targetReflection);

            // Save the workbook to verify the result
            string outputPath = "ReflectionCopyDemo.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="source">The source ReflectionEffect containing the settings to copy.</param>
    /// <param name="destination">The destination ReflectionEffect that will receive the copied settings.</param>
    static void CopyReflectionSettings(ReflectionEffect source, ReflectionEffect destination)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (destination == null) throw new ArgumentNullException(nameof(destination));

        destination.Type = source.Type;
        destination.Blur = source.Blur;
        destination.Distance = source.Distance;
        destination.Direction = source.Direction;
        destination.FadeDirection = source.FadeDirection;
        destination.Size = source.Size;
        destination.Transparency = source.Transparency;
        destination.RotWithShape = source.RotWithShape;
    }
}
