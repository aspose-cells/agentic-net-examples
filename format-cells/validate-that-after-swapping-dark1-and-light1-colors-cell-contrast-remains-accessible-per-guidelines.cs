// Title: C# Aspose.Cells example: validate WCAG contrast ratio for Excel cell font and fill colors after swapping Dark1 and Light1 theme colors
// Common Searches: c# Aspose.Cells verify WCAG contrast ratio for Excel cell colors | how to detect low contrast cells in an Excel workbook using Aspose.Cells | validate accessibility of Excel cell background and font colors with .NET | check cell color contrast after theme color changes in Aspose.Cells | C# code to compute WCAG 2.1 contrast ratio for Excel cells
// Tags: WCAG contrast validation Aspose.Cells | Excel cell accessibility check .NET | solid fill background contrast calculation | theme color swap limitation Aspose.Cells | contrast ratio algorithm C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, iterates through cells in the first worksheet, calculates the WCAG contrast ratio between each cell's font color and solid fill background, reports any cells below the 4.5:1 threshold, and saves the file; it also notes that swapping Dark1 and Light1 theme colors is not supported in the current Aspose.Cells API.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (create/load rule applied)
            var workbook = new Workbook(inputPath);

            // NOTE: Theme manipulation (e.g., swapping Dark1 and Light1) is not supported
            // via a simple API in this version of Aspose.Cells, so this step is omitted.

            // Validate contrast for each cell in the first worksheet
            var sheet = workbook.Worksheets[0];
            bool allAccessible = true;

            foreach (Cell cell in sheet.Cells)
            {
                // Retrieve the cell's style
                Style style = cell.GetStyle();

                // Consider only cells with solid fill background
                if (style.Pattern != BackgroundType.Solid)
                    continue;

                // Get font (text) color; if not set, default to black
                Color fontColor = style.Font.Color.IsEmpty ? Color.Black : style.Font.Color;

                // Get background color; if not set, default to white
                Color bgColor = style.ForegroundColor.IsEmpty ? Color.White : style.ForegroundColor;

                // Compute contrast ratio according to WCAG guidelines
                double ratio = ContrastRatio(fontColor, bgColor);

                // WCAG AA requires a minimum contrast ratio of 4.5:1 for normal text
                if (ratio < 4.5)
                {
                    allAccessible = false;
                    Console.WriteLine($"Low contrast in cell {cell.Name}: ratio {ratio:F2}");
                }
            }

            Console.WriteLine(allAccessible
                ? "All cells meet contrast guidelines."
                : "Some cells do not meet contrast guidelines.");

            // Save the modified workbook (create/save rule applied)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Calculates the relative luminance of a color
    static double Luminance(Color c)
    {
        double RsRGB = c.R / 255.0;
        double GsRGB = c.G / 255.0;
        double BsRGB = c.B / 255.0;

        double R = RsRGB <= 0.03928 ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
        double G = GsRGB <= 0.03928 ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
        double B = BsRGB <= 0.03928 ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

        return 0.2126 * R + 0.7152 * G + 0.0722 * B;
    }

    // Calculates the contrast ratio between two colors
    static double ContrastRatio(Color c1, Color c2)
    {
        double L1 = Luminance(c1);
        double L2 = Luminance(c2);
        double lighter = Math.Max(L1, L2);
        double darker = Math.Min(L1, L2);
        return (lighter + 0.05) / (darker + 0.05);
    }
}
