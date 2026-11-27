// Title: Remove a shape's shadow while preserving reflection and glow effects using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide shape shadow but keep reflection effect | C# remove shadow from Excel shape without affecting glow using Aspose | Set shape shadow transparency to 100% in Aspose.Cells workbook example | Preserve glow effect while disabling shadow on a rectangle shape in Aspose.Cells | How to eliminate shape shadow in Aspose.Cells while retaining other effects
// Tags: shape shadow removal Aspose.Cells C# | preserve reflection effect Aspose.Cells | keep glow effect while hiding shadow Aspose.Cells | Aspose.Cells shape formatting effects | Excel shape shadow transparency .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace RemoveShadowDemo
{
    // This C# example creates a new workbook, adds a rectangle shape, applies reflection and glow effects, then removes the visible shadow by setting its transparency to 100% and resetting size, blur, and distance to zero, ensuring the reflection and glow remain intact before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            const string outputPath = "RemoveShadowDemo.xlsx";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Add a sample shape (rectangle) to demonstrate the effect changes
            // -----------------------------------------------------------------
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = sheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 100);

            // Give the shape a visible fill and line so we can see the effects
            shape.FillFormat.ForeColor = Color.LightBlue;
            shape.LineFormat.Weight = 2;
            shape.LineFormat.ForeColor = Color.DarkBlue;

            // -----------------------------------------------------------------
            // 2. Apply a reflection effect (will be preserved)
            // -----------------------------------------------------------------
            ReflectionEffect reflection = shape.Reflection;
            reflection.Type = ReflectionEffectType.HalfReflection4PtOffset;
            reflection.Transparency = 0.4;   // 40% transparent
            reflection.Size = 55.5;          // size of the reflection
            reflection.Blur = 0.5;           // slight blur
            reflection.Distance = 4;         // distance from the shape

            // -----------------------------------------------------------------
            // 3. Apply a glow effect (will be preserved)
            // -----------------------------------------------------------------
            GlowEffect glow = shape.Glow;
            CellsColor glowColor = workbook.CreateCellsColor();
            glowColor.Color = Color.Yellow;
            glow.Color = glowColor;
            glow.Size = 10;          // radius in points
            glow.Transparency = 0.3; // 30% transparent

            // -----------------------------------------------------------------
            // 4. Remove the shadow effect while keeping reflection and glow
            // -----------------------------------------------------------------
            ShadowEffect shadow = shape.ShadowEffect;

            // The ShadowEffect object does not have a direct "Disable" method.
            // To effectively remove the shadow we make it fully transparent
            // and set its size, blur and distance to zero.
            shadow.Transparency = 1.0; // 100% transparent (invisible)
            shadow.Blur = 0;
            shadow.Distance = 0;
            shadow.Size = 0;

            // If the PresetShadowType enumeration contains a "NoShadow" value,
            // it can be set as well (uncomment the line below if available).
            // shadow.PresetType = PresetShadowType.NoShadow;

            // -----------------------------------------------------------------
            // 5. Save the workbook
            // -----------------------------------------------------------------
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The shape now has reflection and glow, but no visible shadow.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
