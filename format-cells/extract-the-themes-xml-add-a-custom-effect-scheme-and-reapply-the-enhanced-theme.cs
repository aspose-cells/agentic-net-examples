// Title: C# – Apply a Custom 12‑Color Theme and Text‑Effect Shape with Aspose.Cells
// Description: Load a workbook, read its current theme, define a 12‑color palette, create a custom theme named "EnhancedTheme", insert a WordArt shape, switch its text‑effect preset and font settings, then save the file with the new theme and effect scheme using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom theme C# | Workbook.CustomTheme example | apply 12‑color palette Aspose.Cells | add WordArt shape Aspose.Cells | text effect preset C# | change theme programmatically | Aspose.Cells .NET tutorial
// Common Searches: how to create a custom theme in Aspose.Cells .NET | add and modify WordArt text effect with Aspose.Cells | extract current workbook theme Aspose.Cells | apply custom color scheme to Excel file C# | save workbook with custom theme Aspose.Cells
// Developer Intent: Generate a custom 12‑color theme, apply it to a workbook, add a shape with a configurable text‑effect, and save the enhanced spreadsheet.
// Use Cases: Brand corporate reports by defining a reusable color palette as a custom theme. | Highlight sections with decorative WordArt headings that use specific text‑effect presets. | Standardize the appearance of imported workbooks by replacing their default theme with a predefined scheme before distribution.
// AI Prompts: Show C# code that extracts the current theme name and replaces it with a custom 12‑color theme using Aspose.Cells. | Provide an example that adds a WordArt shape, changes its text‑effect preset, and updates font size, bold, and italic properties programmatically. | Explain how to store a custom theme once and reuse it across multiple workbooks without recreating the color array each time.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load a workbook, read its current theme, define a 12‑color palette, create a custom theme named "EnhancedTheme", insert a WordArt shape, switch its text‑effect preset and font settings, then save the file with the new theme and effect scheme using Aspose.Cells for .NET.
class ThemeEnhancer
{
    static void Main()
    {
        // Load an existing workbook (create if not present)
        Workbook workbook = new Workbook("input.xlsx");

        // Extract the current theme name
        string currentTheme = workbook.Theme;
        Console.WriteLine("Current theme: " + currentTheme);

        // Define a custom 12‑color theme (indexes follow Aspose.Cells documentation)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1
            Color.FromArgb(0, 0, 0),       // Text1
            Color.FromArgb(240, 240, 240), // Background2
            Color.FromArgb(80, 80, 80),    // Text2
            Color.FromArgb(255, 0, 0),     // Accent1
            Color.FromArgb(0, 255, 0),     // Accent2
            Color.FromArgb(0, 0, 255),     // Accent3
            Color.FromArgb(255, 255, 0),   // Accent4
            Color.FromArgb(255, 0, 255),   // Accent5
            Color.FromArgb(0, 255, 255),   // Accent6
            Color.FromArgb(0, 0, 255),     // Hyperlink
            Color.FromArgb(255, 0, 0)      // Followed Hyperlink
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("EnhancedTheme", customColors);

        // Add a shape with a custom text effect to demonstrate an effect scheme
        Worksheet sheet = workbook.Worksheets[0];
        Shape wordArt = sheet.Shapes.AddTextEffect(
            MsoPresetTextEffect.TextEffect1, // initial preset
            "Custom Effect",                 // text
            "Calibri",                       // font name
            48,                              // font size
            true,                            // bold
            false,                           // italic
            0, 0, 0, 0,                      // left, top, width, height (auto)
            300, 100);                       // shape width, height

        // Change the text effect to a different preset and adjust properties
        wordArt.TextEffect.SetTextEffect(MsoPresetTextEffect.TextEffect10);
        wordArt.TextEffect.FontBold = true;
        wordArt.TextEffect.FontItalic = false;
        wordArt.TextEffect.FontSize = 36;

        // Save the workbook with the enhanced theme and effect scheme
        workbook.Save("output_enhanced.xlsx");
    }
}
