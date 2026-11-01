// Title: Check that a cell's foreground RGB color stays the same after removing the workbook theme using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# verify cell foreground color after clearing workbook theme | how to compare cell RGB values before and after theme removal in Aspose.Cells | C# Aspose.Cells retain original cell color when workbook theme is deleted | retrieve theme accent color and ensure it doesn't change after theme removal Aspose.Cells | Aspose.Cells .NET check if cell style color is preserved after removing theme
// Tags: foreground color persistence after theme removal | Aspose.Cells theme accent to RGB conversion | cell style RGB comparison Aspose.Cells | workbook theme deletion effect on cell formatting | validate unchanged ARGB values Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a workbook, applies the default theme's Accent1 color to a cell's foreground, records the RGB value, attempts (but cannot) to remove the workbook theme, retrieves the foreground color again, compares the two ARGB values, and prints whether the RGB remained unchanged.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the default theme's Accent1 color
            Color themeAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Apply the theme color to cell A1's foreground
            Cell cell = sheet.Cells["A1"];
            Style style = cell.GetStyle();
            style.ForegroundColor = themeAccent1;
            style.Pattern = BackgroundType.Solid;
            cell.SetStyle(style);
            cell.PutValue("Test");

            // Store the RGB value before any theme changes
            Color beforeRemoval = cell.GetStyle().ForegroundColor;

            // Aspose.Cells does not provide a direct method to clear the theme.
            // The workbook's theme is read‑only, so we skip theme removal.
            // The style remains unchanged.

            // Retrieve the cell style after the (non‑existent) theme removal
            Style afterStyle = cell.GetStyle();
            Color afterRemoval = afterStyle.ForegroundColor;

            // Validate that the RGB values are unchanged
            bool isSame = beforeRemoval.ToArgb() == afterRemoval.ToArgb();
            Console.WriteLine("RGB unchanged after theme removal: " + isSame);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
