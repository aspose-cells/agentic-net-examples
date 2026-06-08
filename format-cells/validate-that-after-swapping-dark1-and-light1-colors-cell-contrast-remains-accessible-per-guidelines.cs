using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeSwapDemo
{
    class Program
    {
        // Calculates relative luminance of a color (per WCAG)
        static double GetLuminance(Color c)
        {
            double R = c.R / 255.0;
            double G = c.G / 255.0;
            double B = c.B / 255.0;

            // Apply sRGB gamma correction
            R = (R <= 0.03928) ? R / 12.92 : Math.Pow((R + 0.055) / 1.055, 2.4);
            G = (G <= 0.03928) ? G / 12.92 : Math.Pow((G + 0.055) / 1.055, 2.4);
            B = (B <= 0.03928) ? B / 12.92 : Math.Pow((B + 0.055) / 1.055, 2.4);

            // Relative luminance
            return 0.2126 * R + 0.7152 * G + 0.0722 * B;
        }

        // Returns contrast ratio between two colors
        static double GetContrastRatio(Color fore, Color back)
        {
            double L1 = GetLuminance(fore);
            double L2 = GetLuminance(back);
            // Ensure L1 is the lighter color
            if (L1 < L2) { double tmp = L1; L1 = L2; L2 = tmp; }
            return (L1 + 0.05) / (L2 + 0.05);
        }

        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Define a style that uses theme colors for background (Dark1) and font (Light1)
            Style themeStyle = workbook.CreateStyle();
            // Background uses ThemeColorType.Background1 (typically Dark1)
            themeStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background1, 0);
            // Font uses ThemeColorType.Text1 (typically Light1)
            themeStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0);
            themeStyle.Font.Size = 12;
            themeStyle.Font.IsBold = true;
            themeStyle.Pattern = BackgroundType.Solid;

            // 3. Apply the style to cell A1
            Cell cell = cells["A1"];
            cell.PutValue("Contrast Test");
            cell.SetStyle(themeStyle);

            // 4. Get the display style (actual colors after theme resolution)
            Style displayBefore = cell.GetDisplayStyle();
            Color foreBefore = displayBefore.Font.Color;
            Color backBefore = displayBefore.BackgroundColor;

            double contrastBefore = GetContrastRatio(foreBefore, backBefore);
            Console.WriteLine($"Contrast before swapping: {contrastBefore:F2} (Accessible: {contrastBefore >= 4.5})");

            // 5. Retrieve current theme colors (12 entries)
            Color[] themeColors = new Color[12];
            for (int i = 0; i < themeColors.Length; i++)
            {
                ThemeColorType type = (ThemeColorType)i;
                themeColors[i] = workbook.GetThemeColor(type);
            }

            // 6. Swap Dark1 (Background1) and Light1 (Text1) colors
            // Background1 = index 0, Text1 = index 1 in the ThemeColorType enum
            Color temp = themeColors[0];
            themeColors[0] = themeColors[1];
            themeColors[1] = temp;

            // 7. Apply the swapped theme to the workbook
            workbook.CustomTheme("SwappedTheme", themeColors);

            // 8. Get the display style again after swapping
            Style displayAfter = cell.GetDisplayStyle();
            Color foreAfter = displayAfter.Font.Color;
            Color backAfter = displayAfter.BackgroundColor;

            double contrastAfter = GetContrastRatio(foreAfter, backAfter);
            Console.WriteLine($"Contrast after swapping: {contrastAfter:F2} (Accessible: {contrastAfter >= 4.5})");

            // 9. Save the workbook for visual verification
            workbook.Save("ThemeSwapContrastDemo.xlsx");
        }
    }
}