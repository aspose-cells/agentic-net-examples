using System;
using System.Drawing;
using Aspose.Cells;

class ThemeResetExample
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Define a fresh set of 12 default theme colors
        // Index mapping:
        // 0-Background1, 1-Text1, 2-Background2, 3-Text2,
        // 4-Accent1, 5-Accent2, 6-Accent3, 7-Accent4,
        // 8-Accent5, 9-Accent6, 10-Hyperlink, 11-Followed Hyperlink
        Color[] freshColors = new Color[]
        {
            Color.White,          // Background1
            Color.Black,          // Text1
            Color.WhiteSmoke,     // Background2
            Color.Gray,           // Text2
            Color.FromArgb(0, 112, 192),   // Accent1 (blue)
            Color.FromArgb(255, 192, 0),   // Accent2 (orange)
            Color.FromArgb(112, 173, 71),  // Accent3 (green)
            Color.FromArgb(255, 0, 0),     // Accent4 (red)
            Color.FromArgb(0, 176, 80),    // Accent5 (lime)
            Color.FromArgb(112, 48, 160),  // Accent6 (purple)
            Color.Blue,           // Hyperlink
            Color.Purple          // Followed Hyperlink
        };

        // Apply the fresh theme to the workbook
        workbook.CustomTheme("FreshDefaultTheme", freshColors);

        // Save the workbook with the new theme
        workbook.Save("OutputWorkbook.xlsx");
    }
}