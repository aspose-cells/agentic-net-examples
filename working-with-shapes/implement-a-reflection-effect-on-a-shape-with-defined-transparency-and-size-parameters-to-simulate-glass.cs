using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class GlassReflectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will act as the "glass" object
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape glassShape = worksheet.Shapes.AddRectangle(2, 1, 2, 1, 150, 100);

                // Access the reflection effect of the shape
                ReflectionEffect reflection = glassShape.Reflection;

                // Configure reflection to simulate a glass‑like appearance
                reflection.Type = ReflectionEffectType.Custom;
                reflection.Transparency = 0.2;   // Low transparency for a subtle start
                reflection.Size = 80;            // Large size to extend the reflection
                reflection.Blur = 10;            // Slight blur for softness
                reflection.Distance = 5;         // Small distance from the shape
                reflection.Direction = 90;       // Reflect vertically downwards
                reflection.FadeDirection = 90;   // Fade in the same direction
                reflection.RotWithShape = true;  // Keep reflection aligned when rotating

                // Optionally rotate the shape to see RotWithShape effect
                glassShape.RotationAngle = 15;

                // Ensure output directory exists before saving
                string outputPath = "GlassReflectionDemo.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the reflection effect applied
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GlassReflectionDemo.Run();
        }
    }
}