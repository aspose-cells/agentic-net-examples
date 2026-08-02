// Title: Change Light1 (Background1) Theme Color to a Subtle Gray while Preserving Accent Colors – Aspose.Cells for .NET
// Description: Creates a new workbook, reads the existing theme palette, replaces the Light1 (Background1) entry with a light‑gray shade that mimics a subtle texture, re‑applies the theme unchanged for all accent colors, demonstrates the new background on cell A1, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | custom Excel theme | Background1 color | Light1 theme | preserve accent colors | subtle gray fill | programmatic workbook styling | Excel theme modification
// Common Searches: Aspose.Cells change Light1 color C# | How to modify Background1 theme without affecting accents | Set subtle gray background in Excel workbook using Aspose | Custom theme with only background color change .NET | Preserve accent palette when updating Excel theme programmatically
// Developer Intent: Generate a custom theme that updates only the Light1 (Background1) color to a subtle gray while leaving every accent color untouched.
// Use Cases: Apply a neutral, low‑contrast background to reports for a printed‑friendly appearance. | Validate theme changes by styling a sample cell after the custom theme is applied. | Distribute Excel files that require a consistent, non‑distracting background across all sheets.
// AI Prompts: Write C# code with Aspose.Cells that replaces the Light1 theme color with a light gray and keeps all other theme colors unchanged. | Show how to read the full theme palette, modify only Background1, and reapply the custom theme in a .NET workbook. | Explain how to create a cell style that uses the new Background1 color and apply it to a range of cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeTextureDemo
{
    // Creates a new workbook, reads the existing theme palette, replaces the Light1 (Background1) entry with a light‑gray shade that mimics a subtle texture, re‑applies the theme unchanged for all accent colors, demonstrates the new background on cell A1, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Retrieve the current theme colors for all types except Background1
            Color[] currentColors = new Color[12];
            for (int i = 1; i < 12; i++)
            {
                // ThemeColorType values map directly to the array index
                ThemeColorType type = (ThemeColorType)i;
                currentColors[i] = workbook.GetThemeColor(type);
            }

            // Define a subtle texture-like color for Light1 (Background1)
            // Since themes accept only solid colors, we use a light gray that mimics a subtle texture
            Color subtleTextureColor = Color.FromArgb(240, 240, 240);
            currentColors[0] = subtleTextureColor; // Background1 (Light1)

            // Apply the custom theme while keeping all accent colors unchanged
            workbook.CustomTheme("CustomWithSubtleTexture", currentColors);

            // (Optional) Apply the new Background1 color to a sample cell to demonstrate the change
            Worksheet sheet = workbook.Worksheets[0];
            Cell sampleCell = sheet.Cells["A1"];
            sampleCell.PutValue("Background1 with subtle texture color");

            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.BackgroundColor = subtleTextureColor; // use the same color for solid fill
            sampleCell.SetStyle(style);

            // Save the workbook (lifecycle rule)
            workbook.Save("Workbook_With_SubtleTexture_Background1.xlsx");
        }
    }
}
