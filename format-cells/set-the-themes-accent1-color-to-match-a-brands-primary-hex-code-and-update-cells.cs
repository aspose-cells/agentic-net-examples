// Title: Set Accent1 Theme Color to a Brand Hex and Style Cells with Aspose.Cells for .NET
// Description: Creates a new workbook, defines the brand primary color #1A73E8, sets the workbook's Accent1 theme color to that hex value, builds a style that uses the theme color for both font and fill, applies the style to cells A1‑A3, and saves the file as BrandThemeAccent1.xlsx.
// Keywords: Aspose.Cells | C# | .NET | theme color | Accent1 | brand hex code | set theme color | style cells | Excel workbook branding | SetThemeColor method
// Common Searches: how to change Accent1 theme color Aspose.Cells .NET | apply brand hex color to Excel cells using Aspose | set workbook theme color programmatically C# | style cells with theme color Aspose.Cells example | save workbook after updating theme colors
// Developer Intent: Update the workbook's Accent1 theme color to a specific brand hex value and use that theme color to format selected cells.
// Use Cases: Brand a spreadsheet by mapping the corporate primary color to Accent1 and applying it to headings or key data cells. | Create a reusable template where changing the Accent1 theme color automatically updates all styled cells across worksheets. | Generate reports that maintain consistent brand colors without hard‑coding RGB values in each style.
// AI Prompts: Show how to set the Accent2 theme color to a different hex value and apply it to a range of cells with Aspose.Cells for .NET. | Provide code that reads a brand color from a JSON configuration file and sets the workbook’s Accent1 theme color accordingly. | Explain how to add a tint or shade to a theme color when styling cells using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeColorDemo
{
    // Creates a new workbook, defines the brand primary color #1A73E8, sets the workbook's Accent1 theme color to that hex value, builds a style that uses the theme color for both font and fill, applies the style to cells A1‑A3, and saves the file as BrandThemeAccent1.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Define the brand's primary color (hex #1A73E8)
            Color brandPrimary = Color.FromArgb(0x1A, 0x73, 0xE8);

            // Set the theme's Accent1 color to the brand's primary color (feature rule: SetThemeColor)
            workbook.SetThemeColor(ThemeColorType.Accent1, brandPrimary);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a style that uses the Accent1 theme color for font and fill
            Style themeStyle = workbook.CreateStyle();
            // Use the theme color for the font (no tint)
            themeStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            // Use the theme color for the cell background (no tint)
            themeStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            themeStyle.Pattern = BackgroundType.Solid;

            // Apply the style to a range of cells
            string[] cellsToStyle = { "A1", "A2", "A3" };
            for (int i = 0; i < cellsToStyle.Length; i++)
            {
                Cell cell = worksheet.Cells[cellsToStyle[i]];
                cell.PutValue($"Styled with Accent1 ({i + 1})");
                cell.SetStyle(themeStyle);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("BrandThemeAccent1.xlsx");
        }
    }
}
