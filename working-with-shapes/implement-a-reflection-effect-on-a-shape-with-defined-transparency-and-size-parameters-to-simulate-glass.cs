// Title: Apply a Glass‑Like Reflection Effect to a Shape using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to create a new Workbook with Aspose.Cells, add a rectangle shape, and configure its ReflectionEffect (type, transparency, size, blur, distance, direction, fade direction, and RotWithShape) to simulate a glass surface. The workbook is saved as GlassReflectionDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape reflection | glass effect | ReflectionEffect | custom reflection | shape transparency | reflection blur | RotWithShape | Excel workbook | add rectangle shape | Aspose.Cells API
// Common Searches: Aspose.Cells how to add reflection to shape | C# set custom reflection properties Aspose.Cells | glass reflection effect Excel shape Aspose | RotWithShape property example Aspose.Cells | ReflectionEffect transparency size C#
// Developer Intent: Generate an Excel file, insert a rectangle, and apply a custom reflection that mimics glass.
// Use Cases: Product catalogs with glossy item thumbnails | Marketing dashboards featuring glass‑like UI elements | Printable reports that need decorative reflective shapes | Interactive Excel templates where shapes rotate while keeping reflections aligned | Educational worksheets demonstrating visual effects
// AI Prompts: Give me C# code to change the reflection size and transparency for a more subtle glass look in Aspose.Cells. | Show how to rotate a shape and keep its reflection using the RotWithShape property. | Explain each ReflectionEffect property and suggest values for a realistic glass appearance. | How can I animate the reflection effect in an Aspose.Cells workbook? | What are the performance considerations when applying custom reflections to many shapes?

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example shows how to create a new Workbook with Aspose.Cells, add a rectangle shape, and configure its ReflectionEffect (type, transparency, size, blur, distance, direction, fade direction, and RotWithShape) to simulate a glass surface. The workbook is saved as GlassReflectionDemo.xlsx.
    public class GlassReflectionDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will act as the glass object
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape glassShape = worksheet.Shapes.AddRectangle(2, 1, 2, 1, 150, 100);

            // Access the reflection effect of the shape
            ReflectionEffect reflection = glassShape.Reflection;

            // Configure reflection to simulate a glass look
            reflection.Type = ReflectionEffectType.Custom;   // Use custom settings
            reflection.Transparency = 0.2;                  // Low transparency for a subtle start
            reflection.Size = 80;                           // Large size to extend the reflection
            reflection.Blur = 15;                           // Soft blur for a smooth fade
            reflection.Distance = 5;                        // Slight offset from the shape
            reflection.Direction = 90;                      // Reflect vertically downwards
            reflection.FadeDirection = 90;                  // Same as direction for consistency
            reflection.RotWithShape = true;                 // Keep reflection aligned when rotating

            // Optional: rotate the shape to see RotWithShape in action
            glassShape.RotationAngle = 10;

            // Save the workbook
            workbook.Save("GlassReflectionDemo.xlsx");
        }
    }
}
