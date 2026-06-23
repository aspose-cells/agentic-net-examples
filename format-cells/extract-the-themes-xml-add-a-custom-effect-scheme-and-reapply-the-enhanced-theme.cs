using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeEnhancement
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // -----------------------------------------------------------------
                // 1. Extract current theme information (the theme name is accessible)
                // -----------------------------------------------------------------
                string originalThemeName = workbook.Theme; // Returns the default theme name
                Console.WriteLine($"Original Theme: {originalThemeName}");

                // -----------------------------------------------------------------
                // 2. Define a custom effect scheme by adding a WordArt shape with a
                //    specific text effect. This demonstrates adding a custom effect
                //    to the workbook.
                // -----------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];

                // Add a WordArt shape (TextEffect) to the worksheet
                // Parameters: preset, text, font name, font size, bold, italic,
                // upperLeftRow, upperLeftColumn, top, left, height, width
                Shape wordArt = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // initial effect (will be changed)
                    "Custom Effect",                 // text
                    "Calibri",                       // font name
                    48,                              // font size
                    false,                           // bold
                    false,                           // italic
                    0, 0,                            // upper left cell
                    0, 0,                            // top & left offset (in points)
                    100, 300);                       // height & width (in points)

                // Modify the text effect to a different preset (e.g., TextEffect10)
                wordArt.TextEffect.SetTextEffect(MsoPresetTextEffect.TextEffect10);
                // Additional formatting for the WordArt
                wordArt.TextEffect.FontBold = true;
                wordArt.TextEffect.FontItalic = false;
                wordArt.TextEffect.FontSize = 36;

                // -----------------------------------------------------------------
                // 3. Create a custom theme by specifying 12 theme colors.
                //    The order of colors follows the Aspose.Cells documentation.
                // -----------------------------------------------------------------
                Color[] customColors = new Color[]
                {
                    Color.FromArgb(255, 255, 255), // Background1 (white)
                    Color.FromArgb(0, 0, 0),       // Text1 (black)
                    Color.FromArgb(240, 240, 240), // Background2 (light gray)
                    Color.FromArgb(80, 80, 80),    // Text2 (dark gray)
                    Color.FromArgb(255, 0, 0),     // Accent1 (red)
                    Color.FromArgb(0, 255, 0),     // Accent2 (green)
                    Color.FromArgb(0, 0, 255),     // Accent3 (blue)
                    Color.FromArgb(255, 165, 0),   // Accent4 (orange)
                    Color.FromArgb(128, 0, 128),   // Accent5 (purple)
                    Color.FromArgb(0, 255, 255),   // Accent6 (cyan)
                    Color.FromArgb(0, 0, 255),     // Hyperlink (blue)
                    Color.FromArgb(255, 0, 0)      // Followed Hyperlink (red)
                };

                // Apply the custom theme (lifecycle rule: method usage)
                workbook.CustomTheme("MyEnhancedTheme", customColors);

                // -----------------------------------------------------------------
                // 4. Demonstrate the new theme by applying an accent color to a cell.
                // -----------------------------------------------------------------
                Cell themedCell = sheet.Cells["A5"];
                themedCell.PutValue("Themed Cell");
                Style themedStyle = workbook.CreateStyle();
                // Use Accent1 from the custom theme
                themedStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                themedStyle.Font.Size = 14;
                themedCell.SetStyle(themedStyle);

                // -----------------------------------------------------------------
                // 5. Save the workbook with the enhanced theme (lifecycle rule: save)
                // -----------------------------------------------------------------
                string outputPath = "EnhancedThemeWorkbook.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with custom theme and effect scheme at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}