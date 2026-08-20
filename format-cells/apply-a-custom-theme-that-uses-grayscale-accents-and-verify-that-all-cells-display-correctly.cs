// Title: Apply a Custom Grayscale Theme in Aspose.Cells for .NET (C#) and Verify Cell Theme Colors
// Description: Creates a new Workbook, defines a 12‑color grayscale palette, applies it with Workbook.CustomTheme, assigns Accent1, Accent2, and Text1 theme colors to cells A1‑A3, prints the ThemeColor types for verification, and saves the file as GrayscaleThemeDemo.xlsx.
// Keywords: Aspose.Cells | C# | custom theme | grayscale theme | Workbook.CustomTheme | theme colors | cell formatting | Excel automation | ThemeColor verification | BackgroundType.Solid
// Common Searches: Aspose.Cells apply custom grayscale theme C# | How to set theme colors for fonts and backgrounds in Aspose.Cells | Verify theme colors in an Aspose.Cells workbook | Workbook.CustomTheme example .NET | Create grayscale Excel reports with Aspose.Cells
// Developer Intent: Generate a grayscale theme, apply it to a workbook, style specific cells with theme colors, and confirm that the colors are correctly applied.
// Use Cases: Produce printable reports that use only grayscale accents for cost‑effective black‑and‑white printing. | Enforce corporate grayscale branding across multiple Excel files generated programmatically. | Create high‑contrast spreadsheets for accessibility where color differentiation is limited.
// AI Prompts: Extend the example to include a custom font family while keeping the grayscale theme. | Write a unit test that asserts the ThemeColor of cells A1, A2, and A3 matches the defined grayscale palette. | Show how to export the workbook to PDF while preserving the grayscale theme colors.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new Workbook, defines a 12‑color grayscale palette, applies it with Workbook.CustomTheme, assigns Accent1, Accent2, and Text1 theme colors to cells A1‑A3, prints the ThemeColor types for verification, and saves the file as GrayscaleThemeDemo.xlsx.
class ApplyGrayscaleTheme
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define 12 grayscale colors for the custom theme
        Color[] grayColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 - white
            Color.FromArgb(0, 0, 0),       // Text1 - black
            Color.FromArgb(240, 240, 240), // Background2 - light gray
            Color.FromArgb(64, 64, 64),    // Text2 - dark gray
            Color.FromArgb(200, 200, 200), // Accent1
            Color.FromArgb(180, 180, 180), // Accent2
            Color.FromArgb(160, 160, 160), // Accent3
            Color.FromArgb(140, 140, 140), // Accent4
            Color.FromArgb(120, 120, 120), // Accent5
            Color.FromArgb(100, 100, 100), // Accent6
            Color.FromArgb(0, 0, 255),     // Hyperlink - blue (kept for visibility)
            Color.FromArgb(128, 0, 128)    // FollowedHyperlink - purple (kept for visibility)
        };

        // Apply the custom grayscale theme
        workbook.CustomTheme("GrayscaleTheme", grayColors);

        // Cell A1: Font uses Accent1 theme color
        Cell cellA1 = sheet.Cells["A1"];
        cellA1.PutValue("Accent1 Font");
        Style styleA1 = workbook.CreateStyle();
        styleA1.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
        cellA1.SetStyle(styleA1);

        // Cell A2: Background uses Accent2 theme color
        Cell cellA2 = sheet.Cells["A2"];
        cellA2.PutValue("Accent2 Background");
        Style styleA2 = workbook.CreateStyle();
        styleA2.Pattern = BackgroundType.Solid;
        styleA2.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
        cellA2.SetStyle(styleA2);

        // Cell A3: Font uses Text1 theme color (should be black)
        Cell cellA3 = sheet.Cells["A3"];
        cellA3.PutValue("Text1 Font");
        Style styleA3 = workbook.CreateStyle();
        styleA3.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0);
        cellA3.SetStyle(styleA3);

        // Verify that the theme colors are applied correctly
        Console.WriteLine("Verification:");
        Console.WriteLine("A1 Font Theme: " + cellA1.GetStyle().Font.ThemeColor.ColorType);
        Console.WriteLine("A2 Background Theme: " + cellA2.GetStyle().ForegroundThemeColor.ColorType);
        Console.WriteLine("A3 Font Theme: " + cellA3.GetStyle().Font.ThemeColor.ColorType);

        // Save the workbook
        workbook.Save("GrayscaleThemeDemo.xlsx");
    }
}
