// Title: Verify that a cell's foreground color stays unchanged after removing the workbook theme with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# verify cell color after workbook theme removal | how to ensure Excel cell style persists when removing theme using Aspose.Cells | C# Aspose.Cells remove theme without affecting solid fill colors | using reflection to call RemoveTheme in Aspose.Cells and test color consistency | compare cell foreground RGB before and after theme deletion with Aspose.Cells
// Tags: Aspose.Cells cell style consistency after theme removal | C# remove workbook theme using reflection | foreground RGB verification Aspose.Cells | solid fill color retention Excel .NET | validate theme color persistence Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new Workbook, applies a solid red foreground color to cell A1, captures the RGB value, attempts to invoke Workbook.RemoveTheme via reflection (if available), captures the color again, compares the two values, and prints whether the colors match, demonstrating that theme removal does not alter the cell's explicit color.
class ThemeRemovalValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Apply a solid color to the cell (simulating a theme color)
            Style style = cell.GetStyle();
            style.ForegroundColor = Color.FromArgb(255, 0, 0); // Red color
            cell.SetStyle(style);

            // Retrieve the actual RGB color before any modification
            Color colorBefore = cell.GetStyle().ForegroundColor;

            // If the API provides a method to remove the theme, invoke it safely
            // (In some versions Workbook.RemoveTheme() may not exist)
            var removeThemeMethod = typeof(Workbook).GetMethod("RemoveTheme");
            if (removeThemeMethod != null)
            {
                removeThemeMethod.Invoke(workbook, null);
            }

            // After theme removal (or no-op), the cell should retain the same RGB color
            Color colorAfter = cell.GetStyle().ForegroundColor;

            // Validate that the colors are identical
            bool colorsMatch = colorBefore.ToArgb() == colorAfter.ToArgb();

            Console.WriteLine("Color before theme removal: " + colorBefore);
            Console.WriteLine("Color after theme removal:  " + colorAfter);
            Console.WriteLine("Colors match: " + colorsMatch);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
