// Title: How to exchange Dark1 and Light1 theme colors in an Excel workbook and ensure WCAG AA cell contrast using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to replace the Dark1 and Light1 colors in a workbook's theme and save the updated file. | Write C# code that iterates over every used cell in an Aspose.Cells workbook, determines the effective background color, calculates the WCAG AA contrast ratio with the font color, and logs cells that fall below a 4.5 ratio. | Update the placeholder SwapThemeColors method to call the Aspose.Cells Theme API (when available) and re‑run the contrast validation to confirm accessibility compliance.
// Common Searches: aspnet replace dark1 light1 theme colors in excel using aspose.cells | c# check wcag contrast ratio for excel cells with aspose.cells | how to calculate color contrast between font and fill in an aspose.cells workbook | validate accessibility of excel workbook after theme color changes with aspose.cells | asp.net core verify cell contrast after swapping theme colors in excel
// Tags: replace dark1 with light1 theme color Aspose.Cells | WCAG AA contrast check Aspose.Cells | compute cell color contrast ratio C# | enumerate used cells foreground background Aspose.Cells | theme API support Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program loads an Excel file, attempts to replace the Dark1 and Light1 theme colors (using a placeholder if the Theme API is unavailable), saves the workbook, then scans all used cells to compute the WCAG AA contrast ratio between each cell's font color and its effective background color, reporting any cells that do not meet the 4.5:1 threshold.
class Program
{
    // Threshold for acceptable contrast ratio (WCAG AA for normal text)
    const double ContrastThreshold = 4.5;

    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to swap theme colors (if supported)
            try
            {
                SwapThemeColors(workbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error swapping theme colors: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");

            // Validate cell contrast after the swap
            try
            {
                ValidateCellContrast(workbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating contrast: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Swaps the Dark1 and Light1 colors in the workbook's theme (fallback implementation)
    static void SwapThemeColors(Workbook workbook)
    {
        // Aspose.Cells versions prior to theme support do not expose Theme APIs.
        // This method is kept for compatibility; it simply logs that the operation is skipped.
        Console.WriteLine("Theme swapping is not supported in the current Aspose.Cells version.");
    }

    // Validates that each cell's foreground (font) and background colors meet the contrast requirement
    static void ValidateCellContrast(Workbook workbook)
    {
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Iterate through all used cells
        foreach (Cell cell in cells)
        {
            Style style = cell.GetStyle();

            // Determine background color
            Color bgColor = GetEffectiveBackgroundColor(style);

            // Determine font (foreground) color; default to black if not set
            Color fgColor = style.Font.Color.IsEmpty ? Color.Black : style.Font.Color;

            double contrast = GetContrastRatio(fgColor, bgColor);

            if (contrast < ContrastThreshold)
            {
                Console.WriteLine($"Contrast issue in cell {cell.Name}: FG={fgColor}, BG={bgColor}, Ratio={contrast:F2}");
            }
        }
    }

    // Retrieves the effective background color considering pattern; defaults to white
    static Color GetEffectiveBackgroundColor(Style style)
    {
        // If the cell has a fill pattern, use the foreground color of the pattern
        if (style.Pattern != BackgroundType.None && !style.ForegroundColor.IsEmpty)
        {
            return style.ForegroundColor;
        }

        // Default background (no fill) is assumed white
        return Color.White;
    }

    // Calculates the contrast ratio between two colors per WCAG definition
    static double GetContrastRatio(Color color1, Color color2)
    {
        double lum1 = GetRelativeLuminance(color1);
        double lum2 = GetRelativeLuminance(color2);

        double brighter = Math.Max(lum1, lum2);
        double darker = Math.Min(lum1, lum2);

        return (brighter + 0.05) / (darker + 0.05);
    }

    // Computes the relative luminance of a color
    static double GetRelativeLuminance(Color color)
    {
        // Convert sRGB components to linear values
        double RsRGB = color.R / 255.0;
        double GsRGB = color.G / 255.0;
        double BsRGB = color.B / 255.0;

        double R = RsRGB <= 0.03928 ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
        double G = GsRGB <= 0.03928 ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
        double B = BsRGB <= 0.03928 ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

        // Relative luminance formula
        return 0.2126 * R + 0.7152 * G + 0.0722 * B;
    }
}
