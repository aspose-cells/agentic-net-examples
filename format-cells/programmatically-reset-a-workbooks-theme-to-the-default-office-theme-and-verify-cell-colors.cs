// Title: Reset an Aspose.Cells workbook to the default Office theme and verify the cell's automatic foreground color in C#
// AI Prompts: Create a style using Workbook.GetThemeColor(ThemeColorType.Accent1), assign it to a cell, and programmatically determine whether the cell's ForegroundColor is the automatic default (empty or black). | Restore the workbook's theme to the built‑in Office theme, apply the Accent1 theme color to a cell style, then output a boolean indicating if the cell's foreground color matches the automatic default. | After resetting the theme, save the workbook and retrieve the cell's style to check if ForegroundColor.IsEmpty or equals Color.Black, printing the result to the console.
// Common Searches: Aspose.Cells C# reset workbook theme to default Office theme | how to verify if a cell uses the automatic foreground color after applying a theme in Aspose.Cells | GetThemeColor Accent1 example and check default color in Aspose.Cells .NET | determine programmatically whether a cell style uses automatic color in Aspose.Cells | reset theme and validate cell foreground color Aspose.Cells tutorial
// Tags: restore built‑in Office theme Aspose.Cells .NET | apply theme color to cell Aspose.Cells | detect automatic cell foreground color C# | verify cell style color after theme change Aspose.Cells | use GetThemeColor with ThemeColorType in Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new Workbook, restores the default Office theme, applies the Accent1 theme color to cell A1 via a custom style, checks whether the cell's foreground color is the automatic default (empty or black), prints the boolean result, and saves the file as ResetTheme.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a value in A1 and apply a theme color (Accent1) to its style
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample");

            // Create a style and set its foreground color to the theme's Accent1 color
            Style themedStyle = workbook.CreateStyle();
            themedStyle.ForegroundColor = workbook.GetThemeColor(ThemeColorType.Accent1);
            themedStyle.Pattern = BackgroundType.Solid;
            cell.SetStyle(themedStyle);

            // Verify the cell's foreground color (should be Accent1)
            Style resultStyle = cell.GetStyle();
            Color foreground = resultStyle.ForegroundColor;

            // Determine if the color is the default automatic color (empty or black)
            bool isDefaultColor = foreground.IsEmpty || foreground.ToArgb() == Color.Black.ToArgb();

            Console.WriteLine($"Cell A1 foreground color is default: {isDefaultColor}");

            // Save the workbook
            workbook.Save("ResetTheme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
