// Title: Aspose.Cells .NET: Remove Shape Shadow While Keeping Reflection and Glow
// Description: Demonstrates how to create a workbook, add a rectangle shape, apply reflection and glow effects, then clear only the shadow effect by resetting its properties. The example saves two files – one with all effects and another with the shadow removed, leaving reflection and glow untouched.
// Keywords: Aspose.Cells | .NET | C# | Excel shape shadow removal | clear shape shadow | preserve reflection effect | preserve glow effect | reset shadow properties | shape visual effects | Aspose.Cells API
// Common Searches: remove shadow from shape Aspose.Cells .NET | keep reflection and glow when clearing shape shadow | reset shadow effect Excel shape using Aspose | Aspose.Cells clear only shadow effect | how to delete shape shadow without affecting other effects
// Developer Intent: The developer needs to delete the shadow effect of a shape while leaving its reflection and glow unchanged.
// Use Cases: Generate a printable Excel report where shadows are omitted but reflection and glow remain for on‑screen versions. | Toggle shadow visibility in a template workbook without altering existing visual styles. | Create two versions of a spreadsheet – one for presentation with full effects and another for export without shadows.
// AI Prompts: Write C# code with Aspose.Cells to clear only the shadow effect of a shape while preserving its reflection and glow. | Explain why resetting shadow properties to default values removes the shadow in Aspose.Cells. | Show how to verify that reflection and glow settings stay the same after the shadow is cleared.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a rectangle shape, apply reflection and glow effects, then clear only the shadow effect by resetting its properties. The example saves two files – one with all effects and another with the shadow removed, leaving reflection and glow untouched.
    public class RemoveShadowPreserveReflectionAndGlow
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // ----- Apply reflection effect (will be preserved) -----
            ReflectionEffect reflection = shape.Reflection;
            reflection.Type = ReflectionEffectType.HalfReflectionTouching;
            reflection.Size = 55;          // percentage
            reflection.Blur = 5;           // points
            reflection.Transparency = 0.3; // 30% transparent

            // ----- Apply glow effect (will be preserved) -----
            GlowEffect glow = shape.Glow;
            glow.Size = 20;                // radius in points
            glow.Color.Color = Color.Yellow;
            glow.Transparency = 0.2;       // 20% transparent

            // ----- Apply shadow effect (to be removed later) -----
            ShadowEffect shadow = shape.ShadowEffect;
            shadow.Size = 1.5;             // size factor
            shadow.Color.Color = Color.Gray;
            shadow.Angle = 135;            // degrees
            shadow.Blur = 10;              // points
            shadow.Transparency = 0.4;     // 40% transparent

            // Save the workbook with all effects applied
            try
            {
                workbook.Save("Shape_WithAllEffects.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook with all effects: {ex.Message}");
                return;
            }

            // ----- Remove only the shadow effect -----
            // Reset shadow properties to defaults (effectively clearing it)
            shadow.Size = 0;
            shadow.Color.Color = Color.Transparent;
            shadow.Angle = 0;
            shadow.Blur = 0;
            shadow.Transparency = 1; // fully transparent
            shadow.Distance = 0;

            // Save the workbook after clearing the shadow effect
            try
            {
                workbook.Save("Shape_ShadowCleared.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook after clearing shadow: {ex.Message}");
            }
        }
    }
}
