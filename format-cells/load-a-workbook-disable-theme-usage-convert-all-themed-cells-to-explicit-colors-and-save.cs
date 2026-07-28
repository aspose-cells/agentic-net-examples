// Title: Aspose.Cells C# – Replace Theme Colors with Concrete RGB Values in an Excel Workbook
// Description: Loads an Excel file, disables theme usage, iterates through each worksheet and populated cell, substitutes ThemeColor references on font, foreground, and background with the actual RGB color obtained via Workbook.GetThemeColor, clears the ThemeColor fields, reapplies the style, and saves the workbook as a new file.
// Keywords: Aspose.Cells | C# | Excel theme colors | GetThemeColor | convert theme to RGB | explicit cell colors | disable theme usage | workbook styling | theme color replacement
// Common Searches: Aspose.Cells convert theme color to RGB | C# replace Excel theme colors with actual colors | remove ThemeColor from cell style Aspose | how to get real color from theme in Aspose.Cells | save workbook without theme references
// Developer Intent: Transform all themed font, foreground, and background colors in an Excel workbook into fixed RGB values using Aspose.Cells for .NET and persist the changes.
// Use Cases: Preserve visual consistency when the file is opened on machines with different Office themes | Prepare workbooks for older Excel formats that do not support theme colors | Guarantee accurate color rendering in PDF or image exports | Enable downstream systems that cannot interpret theme‑based styling
// AI Prompts: Write a C# method that scans every cell in an Aspose.Cells workbook, replaces any ThemeColor with the corresponding RGB from Workbook.GetThemeColor, and clears the ThemeColor property. | Describe how to keep all other style attributes intact while converting theme colors to explicit colors in Aspose.Cells. | Provide code that safely handles empty worksheets and null cell styles during the theme‑color conversion process.

using System;
using System.Drawing;
using Aspose.Cells;

// Loads an Excel file, disables theme usage, iterates through each worksheet and populated cell, substitutes ThemeColor references on font, foreground, and background with the actual RGB color obtained via Workbook.GetThemeColor, clears the ThemeColor fields, reapplies the style, and saves the workbook as a new file.
class ConvertThemedColors
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range; if the sheet is empty, skip it
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;
            if (maxRow < 0 || maxCol < 0) continue;

            // Loop through all cells that contain data or formatting
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell == null) continue;

                    // Get the current style of the cell
                    Style style = cell.GetStyle();

                    // ----- Convert Font Theme Color to explicit color -----
                    ThemeColor fontTheme = style.Font.ThemeColor;
                    if (fontTheme != null && fontTheme.ColorType != ThemeColorType.StyleColor)
                    {
                        // Retrieve the actual theme color from the workbook
                        Color actualColor = workbook.GetThemeColor(fontTheme.ColorType);
                        // Apply the color directly and clear the theme reference
                        style.Font.Color = actualColor;
                        style.Font.ThemeColor = null;
                    }

                    // ----- Convert Foreground Theme Color to explicit color -----
                    ThemeColor fgTheme = style.ForegroundThemeColor;
                    if (fgTheme != null && fgTheme.ColorType != ThemeColorType.StyleColor)
                    {
                        Color actualColor = workbook.GetThemeColor(fgTheme.ColorType);
                        style.ForegroundColor = actualColor;
                        style.ForegroundThemeColor = null;
                    }

                    // ----- Convert Background Theme Color to explicit color -----
                    ThemeColor bgTheme = style.BackgroundThemeColor;
                    if (bgTheme != null && bgTheme.ColorType != ThemeColorType.StyleColor)
                    {
                        Color actualColor = workbook.GetThemeColor(bgTheme.ColorType);
                        style.BackgroundColor = actualColor;
                        style.BackgroundThemeColor = null;
                    }

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
