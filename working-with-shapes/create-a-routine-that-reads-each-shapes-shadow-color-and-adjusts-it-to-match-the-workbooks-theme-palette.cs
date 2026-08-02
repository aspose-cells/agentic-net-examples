// Title: Aspose.Cells .NET – Adjust Shape Shadow Colors to Match Workbook Theme Palette
// Description: A C# routine that loads or creates a workbook, adds a sample shape, then walks through every worksheet and shape, reads each shape's shadow color, finds the nearest theme‑palette color using Workbook.GetMatchingColor, applies the matched color, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | shape shadow color | theme palette | Workbook.GetMatchingColor | Excel shape formatting | color matching | adjust shadow color | Excel automation
// Common Searches: Aspose.Cells change shape shadow to theme color | GetMatchingColor for shape shadow Aspose.Cells | Iterate shapes and update shadow colors .NET | Align shape shadow with workbook theme palette | C# code to normalize Excel shape shadows
// Developer Intent: Replace each shape's shadow color with the closest color from the workbook's theme palette.
// Use Cases: Ensure visual consistency after applying a corporate theme by normalizing all shape shadows. | Convert custom RGB shadow values in imported workbooks to theme‑compatible colors for branding compliance. | Automate cleanup of legacy Excel files that contain non‑theme shadow colors before publishing.
// AI Prompts: Write C# code with Aspose.Cells that iterates all shapes in a workbook and sets each shadow color to the nearest theme palette color. | Show how to safely skip shapes that lack a shadow effect while performing color matching. | Explain Workbook.GetMatchingColor and demonstrate its role in aligning shape shadow colors with the workbook theme.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // A C# routine that loads or creates a workbook, adds a sample shape, then walks through every worksheet and shape, reads each shape's shadow color, finds the nearest theme‑palette color using Workbook.GetMatchingColor, applies the matched color, and saves the file.
    public class AdjustShapeShadowToThemePalette
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // lifecycle: create

                // Example: add a shape with a custom shadow color to demonstrate the adjustment
                Worksheet sheet = workbook.Worksheets[0];
                Shape demoShape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);
                ShadowEffect demoShadow = demoShape.ShadowEffect;
                // Set an arbitrary shadow color (e.g., a custom RGB value)
                demoShadow.Color.Color = Color.FromArgb(123, 200, 150);

                // Iterate through all worksheets
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in ws.Shapes)
                    {
                        // Access the shape's shadow effect
                        ShadowEffect shadow = shape.ShadowEffect;
                        if (shadow == null) continue; // safety check

                        // Retrieve the current shadow color (System.Drawing.Color)
                        Color currentColor = shadow.Color.Color;

                        // Find the closest matching color in the workbook's palette/theme
                        Color matchedColor = workbook.GetMatchingColor(currentColor);

                        // Apply the matched color back to the shadow effect
                        shadow.Color.Color = matchedColor;
                    }
                }

                // Save the workbook (lifecycle: save)
                string outputPath = "AdjustedShadowColors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustShapeShadowToThemePalette.Run();
        }
    }
}
