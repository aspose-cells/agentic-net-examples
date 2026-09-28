// Title: Copy reflection effect settings from one shape to another in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that copies all reflection effect attributes (blur, size, transparency, distance, direction) from a source shape to a target shape in the same worksheet using Aspose.Cells. | Generate a method that clones the reflection effect of one Excel shape onto another shape within an Aspose.Cells workbook. | Provide a script to transfer reflection settings between two shapes in an Excel file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells how to duplicate shape reflection effect in Excel | C# copy reflection properties from one shape to another in a workbook | transfer visual effects between Excel shapes using Aspose.Cells | copy shape reflection settings programmatically with Aspose.Cells .NET | clone shape reflection effect Aspose.Cells example
// Tags: Aspose.Cells shape reflection copy | C# transfer shape visual effect | Excel shape reflection property duplication | reflection effect settings transfer Aspose | clone shape reflection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Copies reflection effect properties (blur, size, transparency, distance, direction) from the first shape to the second shape in the first worksheet of an Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the shape collection
            ShapeCollection shapes = sheet.Shapes;

            // Ensure there are at least two shapes
            if (shapes.Count < 2)
            {
                Console.WriteLine("The worksheet does not contain at least two shapes.");
                return;
            }

            // Identify source and target shapes (first and second)
            Shape sourceShape = shapes[0];
            Shape targetShape = shapes[1];

            // Ensure both shapes have a Reflection effect before copying
            if (sourceShape.Reflection != null && targetShape.Reflection != null)
            {
                // Copy reflection properties from source to target
                targetShape.Reflection.Blur = sourceShape.Reflection.Blur;
                targetShape.Reflection.Size = sourceShape.Reflection.Size;
                targetShape.Reflection.Transparency = sourceShape.Reflection.Transparency;
                targetShape.Reflection.Distance = sourceShape.Reflection.Distance;
                targetShape.Reflection.Direction = sourceShape.Reflection.Direction;
                // Note: RotateWithShape and Alignment are not available in the ReflectionEffect class.
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
