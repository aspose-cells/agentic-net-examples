using System;
using System.Drawing;
using Aspose.Cells;

class ThemeResetExample
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Define a fresh set of 12 default theme colors
        // Index mapping:
        // 0-Background1, 1-Text1, 2-Background2, 3-Text2,
        // 4-Accent1, 5-Accent2, 6-Accent3, 7-Accent4,
        // 8-Accent5, 9-Accent6, 10-Hyperlink, 11-Followed Hyperlink
        Color[] defaultColors = new Color[]
        {
            Color.White,          // Background1
            Color.Black,          // Text1
            Color.FromArgb(242, 242, 242), // Background2 (light gray)
            Color.FromArgb(31, 31, 31),    // Text2 (dark gray)
            Color.FromArgb(0, 112, 192),   // Accent1 (blue)
            Color.FromArgb(255, 192, 0),   // Accent2 (orange)
            Color.FromArgb(112, 173, 71),  // Accent3 (green)
            Color.FromArgb(255, 0, 0),     // Accent4 (red)
            Color.FromArgb(255, 0, 255),   // Accent5 (magenta)
            Color.FromArgb(0, 176, 80),    // Accent6 (teal)
            Color.FromArgb(0, 0, 255),     // Hyperlink (blue)
            Color.FromArgb(128, 0, 128)    // Followed Hyperlink (purple)
        };

        // Apply the fresh theme, effectively removing the previous one
        workbook.CustomTheme("DefaultTheme", defaultColors);

        // Save the workbook with the new theme applied
        workbook.Save("output.xlsx");
    }
}