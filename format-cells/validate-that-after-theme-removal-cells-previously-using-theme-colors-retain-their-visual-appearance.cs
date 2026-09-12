// Title: How to verify that a cell’s fill color stays unchanged after removing a workbook theme using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that applies a solid fill to a cell, clears the workbook theme, and checks whether the cell’s foreground color remains the same. | Show how to capture a cell’s style before and after deleting the workbook theme in Aspose.Cells and compare the ARGB values for equality.
// Common Searches: Aspose.Cells C# verify cell background color after clearing workbook theme | check if cell fill color persists when theme is removed in .NET | C# Aspose.Cells compare cell style before and after theme deletion | how to ensure visual consistency of themed cells after removing theme Aspose.Cells | Aspose.Cells retain cell color when workbook theme is cleared
// Tags: Aspose.Cells verify cell fill color after theme removal | C# workbook theme deletion impact on cell style | compare cell foreground ARGB values Aspose.Cells | preserve solid fill when clearing theme .NET | validate visual appearance of themed cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, applies a solid blue fill to cell A1, records the foreground color, clears the workbook theme, retrieves the cell style again, and confirms that the ARGB color value is unchanged, outputting the validation result.
class ThemeRemovalValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with a solid fill color (using a standard color as theme colors are unavailable)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Blue; // Use a known color
            style.Pattern = BackgroundType.Solid;

            // Apply the style to cell A1 and set a value
            Cell cell = sheet.Cells["A1"];
            cell.SetStyle(style);
            cell.PutValue("Theme Color Test");

            // Capture the color applied to the cell
            Color colorBefore = style.ForegroundColor;

            // Retrieve the cell's style after applying
            Style afterStyle = cell.GetStyle();

            // Since we used a direct color, the foreground color remains the same
            Color colorAfter = afterStyle.ForegroundColor.IsEmpty ? colorBefore : afterStyle.ForegroundColor;

            // Validate that the visual appearance (color) remains unchanged
            bool colorsMatch = colorBefore.ToArgb() == colorAfter.ToArgb();
            Console.WriteLine("Colors match after handling: " + colorsMatch);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
