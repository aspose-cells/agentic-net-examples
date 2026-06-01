using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGlowReader
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains the shape
            Workbook workbook = new Workbook("input.xlsx"); // <-- replace with your file path

            // Access the first worksheet (or any specific worksheet by index/name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Get the shape you want to inspect (e.g., the first shape)
            Shape shape = worksheet.Shapes[0];

            // Access the GlowEffect of the shape
            GlowEffect glow = shape.Glow;

            // Retrieve the color of the glow effect
            CellsColor glowColor = glow.Color;

            // Log the glow color details to the console
            // The Color property returns a System.Drawing.Color instance
            Console.WriteLine("Glow Color Details:");
            Console.WriteLine($"- ARGB: {glowColor.Argb}");
            Console.WriteLine($"- System.Drawing.Color: {glowColor.Color}");
            Console.WriteLine($"- IsShapeColor: {glowColor.IsShapeColor}");
            Console.WriteLine($"- Transparency: {glow.Transparency}");
            Console.WriteLine($"- Size (points): {glow.Size}");

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}