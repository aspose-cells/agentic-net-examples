// Title: Verify cell font color remains unchanged after removing or replacing a workbook theme with Aspose.Cells for .NET
// Description: Shows how to convert a theme‑based font color to a concrete RGB value, replace the workbook theme with the default theme, and confirm that the cell's visual appearance is preserved using Aspose.Cells for .NET.
// Keywords: Aspose.Cells .NET theme color | SetThemeColor Accent1 | CopyTheme reset workbook theme | GetThemeColor RGB conversion | retain cell formatting after theme change | Excel theme removal Aspose | font color comparison before after | theme‑based styling validation | cell style preservation | Aspose.Cells visual consistency
// Common Searches: Aspose.Cells keep cell color after theme removal | convert theme color to RGB Aspose.Cells .NET | copy default theme workbook Aspose | validate formatting after Excel theme reset | how to preserve font color when changing workbook theme
// Developer Intent: Confirm that a cell's visual formatting (font color) is preserved after the workbook's theme is removed or replaced.
// Use Cases: Automated testing to ensure branding colors survive theme resets in generated reports. | Migrating Excel files to older versions while keeping existing theme‑based styling intact. | Creating reusable templates where theme changes must not affect already styled cells.
// AI Prompts: Generate C# code that converts a theme‑based font color to an explicit RGB value and verifies the color after copying the default theme with Aspose.Cells. | Write a unit test in .NET that asserts the font color of cell A1 is identical before and after the workbook theme is replaced. | Explain the steps to preserve cell formatting when swapping an Excel workbook's theme using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeRetentionValidation
{
    // Shows how to convert a theme‑based font color to a concrete RGB value, replace the workbook theme with the default theme, and confirm that the cell's visual appearance is preserved using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // ---------- Step 1: Create workbook and set a custom theme color ----------
            Workbook wb = new Workbook();                                   // create workbook
            wb.SetThemeColor(ThemeColorType.Accent1, Color.Red);           // set Accent1 to Red

            // Apply the theme color to a cell's font
            Worksheet ws = wb.Worksheets[0];
            Cell cell = ws.Cells["A1"];
            cell.PutValue("Theme Color Test");

            Style style = wb.CreateStyle();
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0); // use theme color
            cell.SetStyle(style);

            // Capture the resolved actual color before any theme changes
            Color resolvedBefore = wb.GetThemeColor(ThemeColorType.Accent1);
            Console.WriteLine($"Resolved color before removal: {resolvedBefore}");

            // Save the workbook in its original state
            wb.Save("BeforeRemoval.xlsx");

            // ---------- Step 2: Convert theme‑based color to a concrete RGB color ----------
            // Retrieve the style again, replace the theme reference with the actual color
            Style updatedStyle = cell.GetStyle();
            updatedStyle.Font.Color = resolvedBefore;          // set concrete color
            updatedStyle.Font.ThemeColor = null;               // clear theme reference
            cell.SetStyle(updatedStyle);

            // ---------- Step 3: Remove (replace) the theme ----------
            // Copy the default theme from a fresh workbook, effectively resetting the theme
            Workbook defaultThemeWb = new Workbook();           // default theme workbook
            wb.CopyTheme(defaultThemeWb);                       // replace current theme with default

            // Save the workbook after theme removal
            wb.Save("AfterRemoval.xlsx");

            // ---------- Step 4: Validation ----------
            // Load both workbooks and compare the font colors of cell A1
            Workbook beforeWb = new Workbook("BeforeRemoval.xlsx");
            Workbook afterWb = new Workbook("AfterRemoval.xlsx");

            Color colorBefore = beforeWb.Worksheets[0].Cells["A1"].GetStyle().Font.Color;
            Color colorAfter = afterWb.Worksheets[0].Cells["A1"].GetStyle().Font.Color;

            Console.WriteLine($"Font color before removal: {colorBefore}");
            Console.WriteLine($"Font color after removal : {colorAfter}");

            bool isRetained = colorBefore.ToArgb() == colorAfter.ToArgb();
            Console.WriteLine($"Visual appearance retained after theme removal: {isRetained}");
        }
    }
}
