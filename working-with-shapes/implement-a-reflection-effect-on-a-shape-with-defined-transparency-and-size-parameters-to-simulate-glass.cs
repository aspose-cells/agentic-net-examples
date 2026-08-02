// Title: Aspose.Cells for .NET – C# Example to Add a Custom Glass‑Like Reflection to a Shape
// Description: This example creates a new workbook, inserts a rectangle shape, and applies a custom ReflectionEffect with configurable size, transparency, blur, and distance to mimic a glass surface, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | shape reflection | custom reflection effect | glass appearance | transparency property | blur property | distance property | rectangle shape | Excel workbook example | code snippet
// Common Searches: Aspose.Cells add glass reflection to shape C# | custom reflection effect size transparency Aspose.Cells | how to set blur and distance on shape reflection .NET | Aspose.Cells shape reflection example GitHub | C# code for glass‑like shape effect in Excel
// Developer Intent: Insert a rectangle into a worksheet and configure its ReflectionEffect (type, size, transparency, blur, distance) to achieve a glass‑like look.
// Use Cases: Design a branded spreadsheet header with a reflective glass banner. | Create report elements (callouts, legends) that appear with a polished glass finish. | Build an Excel UI mock‑up where buttons or panels have a realistic reflective surface.
// AI Prompts: Generate code to apply the same custom glass reflection to every shape on a worksheet. | Explain how to adjust the reflection so it fades gradually toward the bottom of the shape. | Show how to export the workbook to PDF while preserving shape reflection effects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, inserts a rectangle shape, and applies a custom ReflectionEffect with configurable size, transparency, blur, and distance to mimic a glass surface, then saves the file as an Excel workbook.
    public class GlassReflectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will represent the glass surface
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 120, 80);

                // Obtain the reflection effect object for the shape
                ReflectionEffect reflection = shape.Reflection;

                // Use a custom reflection type to control all parameters manually
                reflection.Type = ReflectionEffectType.Custom;

                // Set the size of the reflection (percentage of the shape height)
                reflection.Size = 80; // 80% of the shape height

                // Define the starting transparency to achieve a glass‑like look
                reflection.Transparency = 0.3; // 30% transparent (0.0 = opaque, 1.0 = clear)

                // Optional: add a slight blur and distance for a more realistic effect
                reflection.Blur = 5;      // blur radius in points
                reflection.Distance = 2; // distance in points

                // Save the workbook with the applied reflection effect
                workbook.Save("GlassReflectionDemo.xlsx");
                Console.WriteLine("Workbook saved successfully as GlassReflectionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GlassReflectionDemo.Run();
        }
    }
}
