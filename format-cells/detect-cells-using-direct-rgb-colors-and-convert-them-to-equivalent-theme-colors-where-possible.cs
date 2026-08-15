// Title: Aspose.Cells for .NET – Detect Direct RGB Cell Colors and Replace Them with Matching Theme Colors
// Description: This C# example creates a workbook, applies explicit RGB fills and fonts, then scans every worksheet and cell. It identifies foreground or background colors set with a direct RGB value, finds an identical theme color in the workbook’s palette, swaps the RGB for a ThemeColor object, clears the original RGB, and saves the file.
// Keywords: Aspose.Cells RGB to theme color conversion | C# replace Excel cell RGB with theme color | detect direct RGB fill Aspose.Cells | map RGB to Excel theme colors .NET | Aspose.Cells style standardization | Excel theme color matching code
// Common Searches: how to convert cell RGB color to theme color using Aspose.Cells | replace explicit RGB fills with theme colors in C# | find matching Excel theme color for an RGB value | Aspose.Cells scan workbook and change RGB to theme | C# example convert workbook colors to theme palette
// Developer Intent: Programmatically replace cells that use explicit RGB colors with the equivalent theme colors available in the workbook.
// Use Cases: Standardize workbook styling by converting custom RGB fills to the document’s theme palette, ensuring consistent appearance across different Office themes. | Align Excel reports with corporate branding by mapping user‑defined RGB colors to the corresponding theme accent colors. | Decrease file size and simplify maintenance by using reusable theme colors instead of repeated RGB definitions in large spreadsheets.
// AI Prompts: Write a C# method for Aspose.Cells that iterates through all worksheets and converts any cell fill or font color defined with an RGB value to the matching ThemeColorType, then clears the original RGB. | Generate code that logs each cell changed from RGB to a theme color and reports the total number of conversions performed. | Create a reusable utility that accepts a workbook path, performs RGB‑to‑theme conversion as demonstrated, and returns a summary of ThemeColorTypes applied.

using System;
using System.Drawing;
using Aspose.Cells;

// This C# example creates a workbook, applies explicit RGB fills and fonts, then scans every worksheet and cell. It identifies foreground or background colors set with a direct RGB value, finds an identical theme color in the workbook’s palette, swaps the RGB for a ThemeColor object, clears the original RGB, and saves the file.
class ConvertRgbToThemeColors
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample cells with direct RGB colors
        Cell cellA1 = sheet.Cells["A1"];
        cellA1.PutValue("RGB Red");
        Style styleA1 = cellA1.GetStyle();
        styleA1.ForegroundColor = Color.Red;               // Direct RGB
        styleA1.Pattern = BackgroundType.Solid;
        cellA1.SetStyle(styleA1);

        Cell cellA2 = sheet.Cells["A2"];
        cellA2.PutValue("RGB Accent1");
        Style styleA2 = cellA2.GetStyle();
        styleA2.ForegroundColor = workbook.GetThemeColor(ThemeColorType.Accent1); // Same as a theme color
        styleA2.Pattern = BackgroundType.Solid;
        cellA2.SetStyle(styleA2);

        // Iterate through all worksheets and cells to replace RGB colors with theme colors where possible
        foreach (Worksheet ws in workbook.Worksheets)
        {
            int maxRow = ws.Cells.MaxDataRow;
            int maxCol = ws.Cells.MaxDataColumn;

            for (int r = 0; r <= maxRow; r++)
            {
                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = ws.Cells[r, c];
                    Style style = cell.GetStyle();

                    // Convert foreground RGB to theme color
                    if (style.ForegroundColor != Color.Empty && style.ForegroundThemeColor == null)
                    {
                        ThemeColorType? matched = FindMatchingThemeColor(workbook, style.ForegroundColor);
                        if (matched.HasValue)
                        {
                            style.ForegroundThemeColor = new ThemeColor(matched.Value, 0.0);
                            style.ForegroundColor = Color.Empty; // Clear explicit RGB
                        }
                    }

                    // Convert background RGB to theme color
                    if (style.BackgroundColor != Color.Empty && style.BackgroundThemeColor == null)
                    {
                        ThemeColorType? matched = FindMatchingThemeColor(workbook, style.BackgroundColor);
                        if (matched.HasValue)
                        {
                            style.BackgroundThemeColor = new ThemeColor(matched.Value, 0.0);
                            style.BackgroundColor = Color.Empty; // Clear explicit RGB
                        }
                    }

                    cell.SetStyle(style);
                }
            }
        }

        // Save the workbook (lifecycle: save)
        workbook.Save("ConvertedThemeColors.xlsx");
    }

    // Helper: returns a ThemeColorType that exactly matches the given RGB color, or null if none match
    private static ThemeColorType? FindMatchingThemeColor(Workbook wb, Color rgb)
    {
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            Color theme = wb.GetThemeColor(type);
            if (theme.ToArgb() == rgb.ToArgb())
                return type;
        }
        return null;
    }
}
