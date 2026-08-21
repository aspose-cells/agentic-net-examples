// Title: Aspose.Cells .NET – Verify WCAG AA Contrast After Swapping Dark1 and Light1 Theme Colors
// Description: A C# sample that creates a workbook, applies the default theme to cell A1, swaps the Dark1 (Background1) and Light1 (Text1) colors for cell A2, resolves the actual RGB values, computes WCAG 2.0 luminance and contrast ratios, checks the 4.5:1 AA threshold, prints the results, and saves the file for visual review.
// Keywords: Aspose.Cells theme swap | WCAG contrast ratio .NET | Accessibility Excel cells | Dark1 Light1 theme validation | Contrast check Aspose.Cells
// Common Searches: aspocells check contrast after theme swap | wcag 4.5:1 contrast Aspose.Cells .NET | verify accessibility when swapping Excel theme colors | programmatic contrast ratio calculation Aspose.Cells
// Developer Intent: Confirm that exchanging the Dark1 (Background1) and Light1 (Text1) theme colors does not drop cell text/background contrast below the WCAG AA minimum.
// Use Cases: Run an automated pre‑publish audit to ensure theme modifications keep contrast ≥ 4.5:1. | Detect and flag cells that become non‑accessible after applying a custom theme. | Generate a contrast‑ratio report for styled cells in a workbook before distribution.
// AI Prompts: Write a routine that scans all styled cells in an Aspose.Cells workbook, resolves their display colors, calculates WCAG contrast, and lists cells failing a specified threshold. | Create code to swap Background1 and Text1 theme colors for a selected range, then automatically replace any resulting low‑contrast styles with alternative theme colors. | Develop a script that saves both the original and swapped worksheets, then produces a summary table showing contrast ratios for each affected cell.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeSwapDemo
{
    // A C# sample that creates a workbook, applies the default theme to cell A1, swaps the Dark1 (Background1) and Light1 (Text1) colors for cell A2, resolves the actual RGB values, computes WCAG 2.0 luminance and contrast ratios, checks the 4.5:1 AA threshold, prints the results, and saves the file for visual review.
    class Program
    {
        // Calculates relative luminance of a color (per WCAG)
        static double GetLuminance(Color c)
        {
            double RsRGB = c.R / 255.0;
            double GsRGB = c.G / 255.0;
            double BsRGB = c.B / 255.0;

            double R = RsRGB <= 0.03928 ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
            double G = GsRGB <= 0.03928 ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
            double B = BsRGB <= 0.03928 ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

            return 0.2126 * R + 0.7152 * G + 0.0722 * B;
        }

        // Returns contrast ratio between two colors (per WCAG)
        static double GetContrastRatio(Color c1, Color c2)
        {
            double L1 = GetLuminance(c1);
            double L2 = GetLuminance(c2);
            double lighter = Math.Max(L1, L2);
            double darker = Math.Min(L1, L2);
            return (lighter + 0.05) / (darker + 0.05);
        }

        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Cell A1 – original theme usage (Background1 as background, Text1 as foreground)
            Style styleOriginal = workbook.CreateStyle();
            styleOriginal.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background1, 0.0); // Dark1
            styleOriginal.ForegroundThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);      // Light1
            styleOriginal.Pattern = BackgroundType.Solid;
            cells["A1"].PutValue("Original");
            cells["A1"].SetStyle(styleOriginal);

            // Cell A2 – swapped theme usage (Background1 ↔ Text1)
            Style styleSwapped = workbook.CreateStyle();
            styleSwapped.BackgroundThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0); // Light1 becomes background
            styleSwapped.ForegroundThemeColor = new ThemeColor(ThemeColorType.Background1, 0.0); // Dark1 becomes foreground
            styleSwapped.Pattern = BackgroundType.Solid;
            cells["A2"].PutValue("Swapped");
            cells["A2"].SetStyle(styleSwapped);

            // Retrieve the actual display styles (resolved colors)
            Style displayOriginal = cells["A1"].GetDisplayStyle();
            Style displaySwapped = cells["A2"].GetDisplayStyle();

            Color bgOriginal = displayOriginal.BackgroundColor;
            Color fgOriginal = displayOriginal.ForegroundColor;
            Color bgSwapped = displaySwapped.BackgroundColor;
            Color fgSwapped = displaySwapped.ForegroundColor;

            // Compute contrast ratios
            double contrastOriginal = GetContrastRatio(bgOriginal, fgOriginal);
            double contrastSwapped = GetContrastRatio(bgSwapped, fgSwapped);

            // WCAG AA minimum contrast for normal text is 4.5:1
            const double MinimumContrast = 4.5;

            Console.WriteLine($"Original cell contrast ratio: {contrastOriginal:F2}");
            Console.WriteLine($"Swapped cell contrast ratio: {contrastSwapped:F2}");

            Console.WriteLine($"Original cell accessible? {(contrastOriginal >= MinimumContrast ? "Yes" : "No")}");
            Console.WriteLine($"Swapped cell accessible? {(contrastSwapped >= MinimumContrast ? "Yes" : "No")}");

            // Save the workbook for visual verification
            workbook.Save("SwapThemeDemo.xlsx");
        }
    }
}
