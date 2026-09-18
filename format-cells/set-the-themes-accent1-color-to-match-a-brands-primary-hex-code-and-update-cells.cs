// Title: How to set the Accent1 theme color to a brand hex value and style a cell range in Aspose.Cells for .NET
// AI Prompts: Set the workbook's Accent1 theme color to a specific hex code and apply that color as background and font to cells A1:B2 using Aspose.Cells in C#. | Create a custom style from a brand hex color, assign it to ThemeColorType.Accent1, and apply only shading and font color to a defined range with StyleFlag in a .NET workbook.
// Common Searches: Aspose.Cells C# change theme Accent1 color to hex | apply brand color to specific cells using Aspose.Cells .NET | set workbook theme color programmatically Aspose.Cells | use StyleFlag to apply only background and font color in Aspose.Cells | convert hex string to System.Drawing.Color for Aspose.Cells styling
// Tags: set theme accent1 color Aspose.Cells C# | apply brand hex color to cell range Aspose.Cells | create custom style with foreground color Aspose.Cells | use StyleFlag shading fontcolor Aspose.Cells | save workbook after theme update Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, converts a brand hex string to a System.Drawing.Color, sets ThemeColorType.Accent1 to that color, builds a style using the same color for background and font, applies the style to range A1:B2 with shading and font color flags, and saves the workbook as output.xlsx.
class ThemeAccentUpdater
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Brand primary color in hex
            string brandHex = "#1A73E8";

            // Convert hex to Color
            Color brandColor = ColorTranslator.FromHtml(brandHex);

            // Set the theme's Accent1 color to the brand color
            // Use SetThemeColor because Theme property may be a string in some versions
            workbook.SetThemeColor(ThemeColorType.Accent1, brandColor);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style that uses the brand color for background and font
            Style style = workbook.CreateStyle();
            style.ForegroundColor = brandColor;
            style.Pattern = BackgroundType.Solid;
            style.Font.Color = brandColor;

            // Define which cells to update (e.g., A1:B2)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B2");

            // Apply the style to the range (only shading and font color flags are set)
            StyleFlag flag = new StyleFlag
            {
                CellShading = true,
                FontColor = true
            };
            range.ApplyStyle(style, flag);

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
