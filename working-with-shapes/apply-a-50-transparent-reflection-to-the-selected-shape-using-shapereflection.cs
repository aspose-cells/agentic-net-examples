// Title: Apply a 50% Transparent Custom Reflection to a Shape with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a rectangle shape, accesses the read‑only Shape.Reflection object, sets Transparency to 0.5 and Type to Custom, and saves the file. Demonstrates how to use Aspose.Cells Shape.Reflection for semi‑transparent reflections.
// Keywords: Aspose.Cells shape reflection | Shape.Reflection C# | custom reflection effect | shape transparency Aspose.Cells | add rectangle shape .NET | Excel shape reflection example | Aspose.Cells API tutorial
// Common Searches: Aspose.Cells set shape reflection transparency | C# apply 50% transparent reflection to Excel shape | Shape.Reflection custom type Aspose.Cells | how to add rectangle shape with reflection in Aspose.Cells | save workbook after modifying shape reflection
// Developer Intent: Add a rectangle shape to a worksheet and apply a 50% transparent custom reflection using the Shape.Reflection API in Aspose.Cells for .NET.
// Use Cases: Design Excel reports with subtle shape reflections for a polished look. | Generate marketing or branding worksheets where logos appear with semi‑transparent reflections. | Automate creation of visual‑rich workbooks that maintain consistent styling across multiple files.
// AI Prompts: Write C# code that adds a circle shape and applies a 30% transparent reflection using Aspose.Cells. | Show how to change a shape's reflection type to 'Flat' and set transparency to 0.2 in Aspose.Cells. | Provide best‑practice error handling for modifying Shape.Reflection properties in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsReflectionDemo
{
    // C# example that creates a workbook, adds a rectangle shape, accesses the read‑only Shape.Reflection object, sets Transparency to 0.5 and Type to Custom, and saves the file. Demonstrates how to use Aspose.Cells Shape.Reflection for semi‑transparent reflections.
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
                // Parameters: upper left row, upper left column, upper left row offset (pixels),
                // upper left column offset (pixels), width (pixels), height (pixels)
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

                // Obtain the existing ReflectionEffect object from the shape (read‑only property)
                ReflectionEffect reflection = shape.Reflection;

                // Apply a 50% transparent reflection using a custom type
                reflection.Transparency = 0.5; // 0.0 = opaque, 1.0 = fully clear
                reflection.Type = ReflectionEffectType.Custom;

                // Save the workbook with the applied reflection effect
                workbook.Save("ShapeReflection50Percent.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
