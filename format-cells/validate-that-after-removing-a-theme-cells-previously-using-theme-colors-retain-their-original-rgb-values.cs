// Title: Lock cell font color after removing ThemeColor and changing workbook theme – Aspose.Cells for .NET
// Description: Demonstrates how to capture the effective RGB of a theme accent, replace the ThemeColor reference with an explicit Font.Color, and verify that the cell’s color stays unchanged when the workbook’s theme is later modified. Includes custom theme creation, style conversion, and validation logic using Aspose.Cells in C#.
// Keywords: Aspose.Cells theme color to RGB | C# preserve cell color after theme change | convert ThemeColor to explicit color | custom workbook theme Aspose.Cells | lock font color Aspose.Cells | retrieve effective theme RGB | Aspose.Cells .NET theme manipulation | validate color persistence after theme update
// Common Searches: Aspose.Cells keep font color when theme changes | How to convert ThemeColor to RGB in C# | Remove ThemeColor reference Aspose.Cells | Get effective theme accent color Aspose.Cells | Validate cell color after theme update .NET | Aspose.Cells custom theme example | Lock cell color against theme modifications
// Developer Intent: Confirm that a cell’s font color remains unchanged after converting its ThemeColor to an explicit RGB value and then altering the workbook’s theme.
// Use Cases: Capture a theme accent’s RGB value, assign it directly to a cell’s Font.Color, and clear the ThemeColor to protect the visual appearance. | Create a brand‑consistent workbook by applying custom theme colors, then lock specific cells so later theme revisions do not affect them. | Automated testing: assert that cells converted to explicit RGB retain their original color after the workbook’s Accent1 color is updated.
// AI Prompts: Write C# code with Aspose.Cells that converts a cell's ThemeColor to a concrete RGB value, clears the ThemeColor link, changes the workbook theme, and verifies the color stays the same. | Generate a unit test in C# using Aspose.Cells that asserts a cell's font color does not change after the workbook's Accent1 theme color is updated following conversion to explicit RGB. | Explain step‑by‑step how to retrieve the effective RGB of a theme accent, apply it to a cell, remove the ThemeColor reference, and validate persistence after a theme modification.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeValidation
{
    // Demonstrates how to capture the effective RGB of a theme accent, replace the ThemeColor reference with an explicit Font.Color, and verify that the cell’s color stays unchanged when the workbook’s theme is later modified. Includes custom theme creation, style conversion, and validation logic using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a custom theme where Accent1 is Red
            Color[] customColors = new Color[12];
            // Fill with default colors first
            for (int i = 0; i < 12; i++) customColors[i] = Color.White;
            // Set Accent1 (index 4) to Red
            customColors[4] = Color.Red;

            // Apply the custom theme
            workbook.CustomTheme("RedAccentTheme", customColors);

            // Apply the theme color (Accent1) to cell A1 via Font.ThemeColor
            Cell cell = cells["A1"];
            cell.PutValue("Themed Text");
            Style style = workbook.CreateStyle();
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0); // no tint
            cell.SetStyle(style);

            // Capture the effective RGB value before any theme change
            Color originalRgb = workbook.GetThemeColor(ThemeColorType.Accent1);
            Console.WriteLine($"Original RGB from theme (Accent1): {originalRgb}");

            // Convert the cell to use the captured RGB directly (remove theme reference)
            Style updatedStyle = cell.GetStyle();
            updatedStyle.Font.Color = originalRgb;          // set explicit RGB
            updatedStyle.Font.ThemeColor = null;            // clear theme reference
            cell.SetStyle(updatedStyle);

            // Change the theme: set Accent1 to Blue
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Blue);
            Console.WriteLine("Theme Accent1 changed to Blue.");

            // Retrieve the cell's current font color after theme change
            Color afterThemeChangeRgb = cell.GetStyle().Font.Color;
            Console.WriteLine($"Cell RGB after theme change: {afterThemeChangeRgb}");

            // Validate that the RGB value remained unchanged
            bool isUnchanged = afterThemeChangeRgb.ToArgb() == originalRgb.ToArgb();
            Console.WriteLine($"Validation result: {(isUnchanged ? "PASS" : "FAIL")}");

            // Save the workbook (lifecycle save)
            workbook.Save("ThemeValidationResult.xlsx");
        }
    }
}
