// Title: Add Gray12 Texture to Light1 (Background1) Theme Color in Aspose.Cells C#
// Description: Demonstrates how to retrieve the Light1 (Background1) theme color from a workbook, create a Gray12 patterned style that uses the original color as foreground with a light tint as background, apply it to a range (A1:Z100), and save the file—keeping all accent colors unchanged.
// Keywords: Aspose.Cells | C# | theme color texture | Light1 background | Background1 pattern | Gray12 pattern | Excel style pattern | preserve accent colors | apply style to range | Excel workbook theming
// Common Searches: How to add a Gray12 texture to Light1 theme in Aspose.Cells .NET | Replace Background1 color with pattern without changing accents | Apply textured style to a specific range using Aspose.Cells C# | Set theme background pattern in Excel via Aspose.Cells | Create Excel template with patterned Light1 background
// Developer Intent: Apply a subtle Gray12 texture to the Light1 (Background1) theme color while leaving all other theme colors, especially accent colors, untouched.
// Use Cases: Design a report template where the primary background has a professional Gray12 texture. | Highlight a data block (e.g., A1:Z100) with a patterned style without affecting the workbook’s color scheme. | Generate branded Excel files that require a textured Light1 background for visual consistency.
// AI Prompts: Show C# code to apply a Gray12 texture to the Light1 (Background1) theme color in Aspose.Cells while preserving accent colors. | Explain how to retrieve a theme color, create a patterned style, and apply it to a range in an Aspose.Cells workbook. | Provide step‑by‑step instructions for adding a subtle texture to the Light1 background without altering other theme colors.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsThemeTextureDemo
{
    // Demonstrates how to retrieve the Light1 (Background1) theme color from a workbook, create a Gray12 patterned style that uses the original color as foreground with a light tint as background, apply it to a range (A1:Z100), and save the file—keeping all accent colors unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the current Background1 theme color (Light1)
                Color originalBackground = workbook.GetThemeColor(ThemeColorType.Background1);

                // Create a style that uses a subtle texture pattern (Gray12)
                // Foreground color will be the original background color,
                // Background color will be a slightly lighter tint to enhance the texture effect.
                Style textureStyle = workbook.CreateStyle();
                textureStyle.SetPatternColor(
                    BackgroundType.Gray12,                 // Subtle texture pattern
                    originalBackground,                    // Foreground color (original Light1)
                    Color.FromArgb(255, 245, 245, 245));    // Very light background for contrast

                // Apply the style to a defined range (A1:Z100)
                AsposeRange range = sheet.Cells.CreateRange("A1:Z100");
                range.ApplyStyle(textureStyle, new StyleFlag { All = true });

                // Save the workbook
                workbook.Save("ThemeWithTexture.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
