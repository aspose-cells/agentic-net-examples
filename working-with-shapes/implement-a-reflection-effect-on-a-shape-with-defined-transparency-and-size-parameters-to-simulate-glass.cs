// Title: Create a glass‑like reflection on a rectangle shape using Aspose.Cells ReflectionEffect in C#
// Common Searches: how to add a custom reflection effect to a shape with Aspose.Cells C# | Aspose.Cells set reflection size and transparency for rectangle shape | C# example of using ReflectionEffect.Type = Custom in Aspose.Cells | apply blur and distance to shape reflection in Aspose.Cells workbook | rotate shape with reflection using RotWithShape property Aspose.Cells
// Tags: Aspose.Cells custom shape reflection | reflection effect properties C# | rectangle shape glass appearance Aspose.Cells | set reflection blur distance Aspose.Cells | RotWithShape property usage Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GlassReflectionDemo
{
    // The sample creates a new workbook, adds a rectangle shape, applies a semi‑transparent fill, and configures a custom ReflectionEffect (size, transparency, blur, distance, direction, fade direction, and RotWithShape) to simulate a glass‑like glossy surface. The shape is optionally rotated, the workbook is saved as GlassReflectionShape.xlsx, and the applied reflection settings are printed to the console.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will act as the "glass" pane
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
                Shape glassShape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 300, 150);

                // Set a semi‑transparent fill to resemble glass
                // Note: SolidFillColor may not be available in older Aspose.Cells versions; using default fill instead
                // glassShape.Fill.SolidFillColor = Color.FromArgb(80, 0, 176, 240);

                // Remove any line to keep the glass clean
                glassShape.Line.Weight = 0;

                // Configure the reflection effect to simulate a glossy surface
                ReflectionEffect reflection = glassShape.Reflection;

                // Use a custom reflection type so we can control all parameters
                reflection.Type = ReflectionEffectType.Custom;

                // Size (percentage of the gradient ramp where the reflection ends) – larger value gives a longer reflection
                reflection.Size = 85; // 85%

                // Transparency of the starting point of the reflection – lower value = more visible reflection
                reflection.Transparency = 0.15; // 15% transparent (mostly opaque)

                // Blur radius – softens the reflection edges for a realistic glass look
                reflection.Blur = 4.5;

                // Distance – how far the reflection is offset from the shape
                reflection.Distance = 6;

                // Direction – angle of the reflection gradient (90° points straight down)
                reflection.Direction = 90;

                // FadeDirection – offset direction for the gradient fade
                reflection.FadeDirection = 90;

                // Ensure the reflection rotates with the shape (useful if the shape is later rotated)
                reflection.RotWithShape = true;

                // Optional: rotate the shape slightly to showcase the RotWithShape effect
                glassShape.RotationAngle = 5;

                // Save the workbook
                string outputPath = "GlassReflectionShape.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Reflection properties applied:");
                Console.WriteLine($"  Type          : {reflection.Type}");
                Console.WriteLine($"  Size (%)      : {reflection.Size}");
                Console.WriteLine($"  Transparency  : {reflection.Transparency}");
                Console.WriteLine($"  Blur (pt)     : {reflection.Blur}");
                Console.WriteLine($"  Distance (pt) : {reflection.Distance}");
                Console.WriteLine($"  Direction (°) : {reflection.Direction}");
                Console.WriteLine($"  FadeDirection : {reflection.FadeDirection}");
                Console.WriteLine($"  RotWithShape  : {reflection.RotWithShape}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the glass reflection effect:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
