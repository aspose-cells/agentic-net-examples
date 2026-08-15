// Title: Aspose.Cells C# – Set Workbook Accent1 Theme Color from a Brand Hex Code and Apply to Cell Font & Fill
// Description: Demonstrates how to convert a brand hex value (e.g., #1A73E8) to a System.Drawing.Color, assign it to the workbook's Accent1 theme slot with SetThemeColor, and style a cell (A1) to use that theme color for both font and background before saving the file.
// Keywords: Aspose.Cells | C# | .NET | SetThemeColor | ThemeColorType.Accent1 | brand hex color | Excel theme color programmatically | cell font color | cell background fill | Workbook theming | Excel styling example
// Common Searches: how to set Accent1 theme color in Aspose.Cells C# | apply brand hex code to Excel theme using Aspose.Cells | change workbook theme color programmatically .NET | use theme color for cell font and fill Aspose.Cells | SetThemeColor example C#
// Developer Intent: Set the workbook’s Accent1 theme color to a specific brand hex value and use that theme color for cell font and background formatting.
// Use Cases: Create corporate‑branded reports where headings automatically inherit the brand’s primary color. | Generate multiple workbooks with a consistent visual identity by updating Accent1 once per file. | Read a brand color from a configuration file or database and apply it to the workbook theme before populating data.
// AI Prompts: Show C# code to adjust the tint of the Accent1 theme color after calling SetThemeColor in Aspose.Cells. | Explain how to revert a workbook’s custom theme back to the default theme using Aspose.Cells for .NET. | Provide a method to read the current Accent1 theme color from an existing workbook and replace it with a new hex value.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to convert a brand hex value (e.g., #1A73E8) to a System.Drawing.Color, assign it to the workbook's Accent1 theme slot with SetThemeColor, and style a cell (A1) to use that theme color for both font and background before saving the file.
class ThemeAccentDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Brand's primary color in hex (example: #1A73E8)
        Color brandPrimary = ColorTranslator.FromHtml("#1A73E8");

        // Set the theme's Accent1 color to the brand's primary color
        workbook.SetThemeColor(ThemeColorType.Accent1, brandPrimary);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Apply the Accent1 theme color to a cell's font and background
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Brand Accent1 Theme");

        // Retrieve the cell's current style
        Style style = cell.GetStyle();

        // Use the Accent1 theme color for the font (no tint)
        style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);

        // Use the same theme color for the cell background
        style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
        style.Pattern = BackgroundType.Solid;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("BrandTheme.xlsx");
    }
}
