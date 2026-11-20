// Title: How to verify WCAG 2.0 contrast ratio after swapping Dark1 and Light1 theme colors in an Excel cell using Aspose.Cells for .NET
// Common Searches: calculate WCAG 2.0 contrast ratio between cell background and font colors with Aspose.Cells C# | swap Dark1 and Light1 theme colors in Aspose.Cells and check accessibility compliance | programmatically validate Excel cell color contrast using Aspose.Cells .NET API | determine if cell text meets 4.5:1 contrast requirement after theme color swap in Aspose.Cells | C# example for computing relative luminance of colors in an Aspose.Cells workbook
// Tags: WCAG contrast ratio calculation Aspose.Cells | swap theme colors Dark1 Light1 Aspose.Cells | cell style accessibility validation .NET | relative luminance method C# Aspose.Cells | Excel workbook color contrast check Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, swaps simulated Dark1 and Light1 theme colors, applies the swapped colors to a cell's background and font, computes the WCAG 2.0 contrast ratio using custom luminance functions, reports whether the 4.5:1 accessibility threshold is met, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Define two colors to simulate theme colors (Dark1 and Light1)
            Color darkColor = Color.FromArgb(0, 0, 0);          // Black (Dark1)
            Color lightColor = Color.FromArgb(255, 255, 255); // White (Light1)

            // Swap the colors
            Color temp = darkColor;
            darkColor = lightColor;
            lightColor = temp;

            // Prepare a sample cell that uses the swapped colors
            var sheet = workbook.Worksheets[0];
            var cell = sheet.Cells["B2"];
            var style = cell.GetStyle();

            // Use swapped Dark1 as background and Light1 as font color
            style.ForegroundColor = darkColor; // background
            style.Pattern = BackgroundType.Solid;
            style.Font.Color = lightColor;     // font

            cell.SetStyle(style);
            cell.PutValue("Contrast Test");

            // Validate contrast ratio between background and font colors
            double contrast = GetContrastRatio(style.ForegroundColor, style.Font.Color);
            Console.WriteLine($"Contrast ratio: {contrast:F2}");

            if (contrast >= 4.5)
                Console.WriteLine("Contrast meets accessibility guidelines.");
            else
                Console.WriteLine("Contrast does NOT meet accessibility guidelines.");

            // Save the workbook (optional)
            string outputPath = "SwappedTheme.xlsx";
            try
            {
                // Ensure the directory exists (handle cases where outputPath has no directory)
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Calculates WCAG 2.0 contrast ratio between two colors
    static double GetContrastRatio(Color bg, Color fg)
    {
        double lumBg = GetRelativeLuminance(bg);
        double lumFg = GetRelativeLuminance(fg);

        // Ensure lumBg is the lighter color
        if (lumBg < lumFg)
        {
            var temp = lumBg;
            lumBg = lumFg;
            lumFg = temp;
        }

        return (lumBg + 0.05) / (lumFg + 0.05);
    }

    // Computes relative luminance of a color per WCAG formula
    static double GetRelativeLuminance(Color c)
    {
        double RsRGB = c.R / 255.0;
        double GsRGB = c.G / 255.0;
        double BsRGB = c.B / 255.0;

        double R = RsRGB <= 0.03928 ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
        double G = GsRGB <= 0.03928 ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
        double B = BsRGB <= 0.03928 ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

        return 0.2126 * R + 0.7152 * G + 0.0722 * B;
    }
}
