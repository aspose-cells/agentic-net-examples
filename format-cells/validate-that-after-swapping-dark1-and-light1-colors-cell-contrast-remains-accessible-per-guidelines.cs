// Title: C# – Verify Accessibility Contrast After Swapping Dark1 and Light1 Theme Colors with Aspose.Cells
// Description: This example creates an Excel workbook, applies the Dark1 (Background1) and Light1 (Text1) theme colors to two cells, swaps the background and font theme colors, resolves the actual RGB values with GetDisplayStyle, checks that foreground and background colors remain distinct, logs the result, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | theme colors | Dark1 | Light1 | swap theme colors | GetDisplayStyle | cell contrast | accessibility | WCAG | Excel automation | color validation | Workbook styling
// Common Searches: how to swap Dark1 and Light1 in Aspose.Cells | validate cell contrast after theme change Aspose.Cells .NET | Aspose.Cells GetDisplayStyle example | C# check Excel cell accessibility contrast | programmatically ensure WCAG contrast in generated Excel files
// Developer Intent: Confirm that exchanging Dark1 and Light1 theme colors does not break the required foreground‑background contrast for accessibility.
// Use Cases: Generate reports that dynamically switch theme colors while preserving WCAG‑compliant contrast. | Automate accessibility audits of Excel files by validating color contrast after style modifications. | Create templates where theme swaps are applied to specific cells and the contrast outcome is logged for quality control.
// AI Prompts: Write a C# function using Aspose.Cells that calculates the WCAG contrast ratio between a cell's resolved background and font colors. | Provide code to iterate over a range, swap BackgroundThemeColor and Font.ThemeColor for each cell, and collect cells that fail a contrast threshold. | Explain how GetDisplayStyle resolves theme colors in Aspose.Cells and how it can be used for accessibility validation.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeSwapDemo
{
    // This example creates an Excel workbook, applies the Dark1 (Background1) and Light1 (Text1) theme colors to two cells, swaps the background and font theme colors, resolves the actual RGB values with GetDisplayStyle, checks that foreground and background colors remain distinct, logs the result, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Prepare two cells with original Dark1 (Background1) and Light1 (Text1) theme colors
            // ------------------------------------------------------------

            // Cell A1: Dark background, Light foreground
            Style styleA1 = workbook.CreateStyle();
            styleA1.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background1, 0.0); // Dark1
            styleA1.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);            // Light1
            styleA1.Pattern = BackgroundType.Solid;
            cells["A1"].PutValue("Original");
            cells["A1"].SetStyle(styleA1);

            // Cell A2: Light background, Dark foreground
            Style styleA2 = workbook.CreateStyle();
            styleA2.BackgroundThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);      // Light1
            styleA2.Font.ThemeColor = new ThemeColor(ThemeColorType.Background1, 0.0);    // Dark1
            styleA2.Pattern = BackgroundType.Solid;
            cells["A2"].PutValue("Original");
            cells["A2"].SetStyle(styleA2);

            // ------------------------------------------------------------
            // 2. Swap Dark1 and Light1 theme colors for both cells
            // ------------------------------------------------------------
            SwapThemeColors(cells["A1"]);
            SwapThemeColors(cells["A2"]);

            // ------------------------------------------------------------
            // 3. Validate that foreground and background colors still provide contrast
            // ------------------------------------------------------------
            ValidateContrast(cells["A1"]);
            ValidateContrast(cells["A2"]);

            // ------------------------------------------------------------
            // 4. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("ThemeSwapContrastDemo.xlsx", SaveFormat.Xlsx);
        }

        // Swaps the background and font theme colors of a given cell
        private static void SwapThemeColors(Cell cell)
        {
            // Retrieve the current style
            Style curStyle = cell.GetStyle();

            // Store original theme colors
            ThemeColor bgTheme = curStyle.BackgroundThemeColor;
            ThemeColor fontTheme = curStyle.Font.ThemeColor;

            // Swap them
            curStyle.BackgroundThemeColor = fontTheme;
            curStyle.Font.ThemeColor = bgTheme;

            // Apply the modified style back to the cell
            cell.SetStyle(curStyle);
        }

        // Checks that the effective foreground and background colors are not identical
        private static void ValidateContrast(Cell cell)
        {
            // Get the display style (takes theme resolution into account)
            Style displayStyle = cell.GetDisplayStyle();

            // Resolve actual colors
            Color bgColor = displayStyle.BackgroundColor;
            Color fgColor = displayStyle.Font.Color;

            // Simple contrast check: colors must differ
            bool hasContrast = bgColor.ToArgb() != fgColor.ToArgb();

            Console.WriteLine($"Cell {cell.Name}: Background={bgColor}, Foreground={fgColor}, ContrastOK={hasContrast}");
        }
    }
}
