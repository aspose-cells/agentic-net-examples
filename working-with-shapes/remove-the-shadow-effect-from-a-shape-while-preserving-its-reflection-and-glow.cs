using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class RemoveShadowPreserveReflectionGlow
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

                // ----- Apply initial effects -----
                // Glow effect
                GlowEffect glow = shape.Glow;
                glow.Size = 20;
                glow.Color.Color = Color.Yellow;
                glow.Transparency = 0.3;

                // Reflection effect
                ReflectionEffect reflection = shape.Reflection;
                reflection.Type = ReflectionEffectType.HalfReflectionTouching;
                reflection.Size = 60;
                reflection.Blur = 5;
                reflection.Transparency = 0.4;
                reflection.Distance = 10;

                // Shadow effect (to be removed later)
                ShadowEffect shadow = shape.ShadowEffect;
                shadow.Size = 1.0;
                shadow.Color.Color = Color.Gray;
                shadow.Angle = 135;
                shadow.Distance = 5;
                shadow.Blur = 10;
                shadow.Transparency = 0.5;

                // Save workbook with all effects applied
                workbook.Save("Shape_WithAllEffects.xlsx");

                // ----- Remove only the shadow effect -----
                // Aspose.Cells does not provide a direct ClearShadowEffect method.
                // Setting the shadow's transparency to 1 (fully transparent) effectively removes it
                // while keeping glow and reflection intact.
                shape.ShadowEffect.Transparency = 1.0;
                shape.ShadowEffect.Size = 0;

                // Save workbook after shadow removal
                workbook.Save("Shape_ShadowRemoved.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            RemoveShadowPreserveReflectionGlow.Run();
        }
    }
}