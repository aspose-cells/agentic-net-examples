// Title: Validate cell color persistence after theme alteration with Aspose.Cells for .NET
// Description: Shows how to capture a theme‑based color, modify the workbook theme, replace the theme reference with the original RGB value, and confirm that the cell's visual appearance stays the same. The workbook is saved as an .xlsx file for manual verification.
// Keywords: Aspose.Cells | C# theme color | theme removal | flatten theme colors | GetThemeColor | SetThemeColor | explicit RGB conversion | Excel theme conversion | cell style preservation | Aspose.Cells .NET
// Common Searches: How to keep cell background color after changing Excel theme in Aspose.Cells | Aspose.Cells replace theme color with RGB | Validate visual consistency after theme change C# | Convert themed cells to explicit colors Aspose.Cells | Preserve cell formatting when removing workbook theme
// Developer Intent: Ensure that a cell styled with a theme color displays the identical visual color after the workbook's theme is altered or removed.
// Use Cases: Capture the original theme color, change the theme, then substitute the theme reference with the captured RGB to keep the look unchanged. | Batch‑process a worksheet to flatten all theme‑based styles into explicit RGB values for compatibility with older Excel versions or third‑party viewers. | Create an automated test that asserts visual consistency of themed cells after programmatically updating the workbook theme.
// AI Prompts: Write C# code using Aspose.Cells that replaces every theme‑based cell style in a worksheet with its resolved RGB color and verifies that the appearance does not change. | Explain how GetThemeColor and SetThemeColor can be combined to test the effect of theme removal on cell formatting in Aspose.Cells. | Generate a unit test in C# that asserts a cell's foreground color remains identical after the workbook's Accent1 theme color is changed.

using Aspose.Cells;
using System;
using System.Drawing;

// Shows how to capture a theme‑based color, modify the workbook theme, replace the theme reference with the original RGB value, and confirm that the cell's visual appearance stays the same. The workbook is saved as an .xlsx file for manual verification.
class ThemeRemovalValidation
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Apply a theme color (Accent1) to cell A1
        Cell cell = ws.Cells["A1"];
        cell.PutValue("Theme Color Cell");
        Style style = wb.CreateStyle();
        style.Pattern = BackgroundType.Solid;
        style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
        cell.SetStyle(style);

        // Capture the resolved color from the current theme (before any change)
        Color resolvedBefore = wb.GetThemeColor(ThemeColorType.Accent1);
        Console.WriteLine("Resolved theme color before change: " + resolvedBefore);

        // Change the theme color to a different color (simulating theme removal)
        wb.SetThemeColor(ThemeColorType.Accent1, Color.Red);
        Color resolvedAfter = wb.GetThemeColor(ThemeColorType.Accent1);
        Console.WriteLine("Resolved theme color after change: " + resolvedAfter);

        // Replace the theme reference with the explicit color captured earlier
        Style updatedStyle = cell.GetStyle();
        updatedStyle.ForegroundColor = resolvedBefore; // set explicit RGB color
        cell.SetStyle(updatedStyle);

        // Verify that the cell's displayed color matches the original resolved color
        Color finalColor = cell.GetStyle().ForegroundColor;
        Console.WriteLine("Final cell foreground color: " + finalColor);

        // Save the workbook for visual inspection
        wb.Save("ThemeRemovalValidation.xlsx");
    }
}
