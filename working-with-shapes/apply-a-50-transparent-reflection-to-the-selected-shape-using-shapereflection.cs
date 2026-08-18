// Title: Apply 50% Transparent Custom Reflection to a Shape with Aspose.Cells for .NET
// Description: Creates a new workbook, adds a rectangle shape, accesses its Reflection property, sets ReflectionEffect.Type to Custom, configures Transparency (0.5), Size (55), Blur (0.5) and Distance (0), then saves the file as ShapeWithReflection.xlsx.
// Keywords: Aspose.Cells shape reflection | C# ReflectionEffect | custom shape reflection .NET | shape transparency Aspose.Cells | reflection blur size distance
// Common Searches: Aspose.Cells add transparent reflection to shape | C# set custom reflection properties on worksheet shape | how to use Shape.Reflection in Aspose.Cells | save workbook with reflected rectangle shape
// Developer Intent: Add a 50% transparent custom reflection to a rectangle shape and persist the workbook.
// Use Cases: Highlight key totals in financial dashboards with reflective rectangles. | Generate marketing PDFs where shapes have a subtle semi‑transparent shine. | Apply a uniform reflection style to multiple shapes across automated reports.
// AI Prompts: Show C# code to change the reflection size and blur of an existing shape in Aspose.Cells. | Generate a script that applies a 50% transparent custom reflection to every shape in a workbook. | Explain how each ReflectionEffect property (Type, Transparency, Size, Blur, Distance) influences the visual output.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape, accesses its Reflection property, sets ReflectionEffect.Type to Custom, configures Transparency (0.5), Size (55), Blur (0.5) and Distance (0), then saves the file as ShapeWithReflection.xlsx.
class ApplyReflection
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

            // Obtain the reflection effect from the shape
            ReflectionEffect reflection = shape.Reflection;

            // Set custom reflection type and its properties
            reflection.Type = ReflectionEffectType.Custom;
            reflection.Transparency = 0.5;   // 50% transparency
            reflection.Size = 55;           // end alpha position (percentage)
            reflection.Blur = 0.5;          // blur radius
            reflection.Distance = 0;       // distance from the shape

            // Save the workbook with the applied reflection effect
            workbook.Save("ShapeWithReflection.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
