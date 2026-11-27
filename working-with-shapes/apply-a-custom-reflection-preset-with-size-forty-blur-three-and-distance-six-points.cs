// Title: Create a rectangle shape and apply a custom reflection preset (size 40%, blur 3pt, distance 6pt) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom reflection size blur distance on a shape | how to add rectangle shape with reflection effect in Aspose.Cells workbook | C# example for configuring shape reflection properties in Aspose.Cells | set reflection transparency and custom preset for worksheet shape using Aspose.Cells | apply custom reflection effect to shape in Aspose.Cells for .NET
// Tags: Aspose.Cells shape reflection custom preset | C# rectangle shape reflection effect | set reflection size blur distance Aspose.Cells | worksheet shape reflection transparency | Aspose.Cells workbook shape styling

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a rectangle shape on the first worksheet, switches the shape's ReflectionEffect to a custom preset, and configures the reflection size to 40%, blur radius to 3 points, distance to 6 points, with full opacity, then saves the file as CustomReflectionDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);

            // Access the reflection effect of the shape
            ReflectionEffect reflection = shape.Reflection;

            // Use a custom reflection preset
            reflection.Type = ReflectionEffectType.Custom;

            // Apply the requested custom settings
            reflection.Size = 40;      // Size in percentage
            reflection.Blur = 3;       // Blur radius in points
            reflection.Distance = 6;   // Distance in points

            // Optional: ensure the reflection is fully opaque (default is 0)
            reflection.Transparency = 0.0;

            // Save the workbook
            string outputPath = "CustomReflectionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
