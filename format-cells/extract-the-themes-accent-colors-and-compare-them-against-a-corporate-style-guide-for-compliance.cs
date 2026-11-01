// Title: Check Excel workbook theme accent colors for corporate style guide compliance using Aspose.Cells for .NET (C#)
// Common Searches: how to read Excel theme accent colors with Aspose.Cells in C# | C# code to compare workbook theme colors to a corporate color palette | validate Excel theme palette against corporate style guide using Aspose.Cells | Aspose.Cells GetThemeColor example for theme compliance check | detect non‑compliant theme colors in an .xlsx file with C#
// Tags: Aspose.Cells GetThemeColor API | extract Excel theme accent palette C# | corporate style guide color matching | theme compliance check Aspose.Cells | validate workbook theme colors .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook, uses Workbook.GetThemeColor with ThemeColorType.Accent1‑Accent6 to retrieve the six theme accent colors, defines a corporate color palette, compares each extracted color's ARGB value to the corporate set, logs matches or mismatches, and reports whether the workbook's theme fully complies with the corporate style guide.
class ThemeAccentCompliance
{
    static void Main()
    {
        try
        {
            string filePath = "InputWorkbook.xlsx";

            // Ensure the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Extract the six accent colors from the workbook theme
            Color[] themeAccents = new Color[6];
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            for (int i = 0; i < accentTypes.Length; i++)
            {
                themeAccents[i] = workbook.GetThemeColor(accentTypes[i]);
            }

            // Corporate style guide accent colors (example values)
            Color[] corporateAccents = new Color[]
            {
                Color.FromArgb(0, 120, 215),   // Corporate Accent 1
                Color.FromArgb(232, 17, 35),   // Corporate Accent 2
                Color.FromArgb(255, 185, 0),   // Corporate Accent 3
                Color.FromArgb(0, 204, 106),   // Corporate Accent 4
                Color.FromArgb(255, 0, 255),   // Corporate Accent 5
                Color.FromArgb(0, 0, 0)        // Corporate Accent 6
            };

            // Compare theme accents against corporate accents
            bool allCompliant = true;
            for (int i = 0; i < themeAccents.Length; i++)
            {
                bool match = false;
                foreach (Color corpColor in corporateAccents)
                {
                    if (themeAccents[i].ToArgb() == corpColor.ToArgb())
                    {
                        match = true;
                        break;
                    }
                }

                if (!match)
                {
                    allCompliant = false;
                    Console.WriteLine($"Accent {i + 1} (#{themeAccents[i].ToArgb():X8}) does NOT match corporate style guide.");
                }
                else
                {
                    Console.WriteLine($"Accent {i + 1} (#{themeAccents[i].ToArgb():X8}) matches corporate style guide.");
                }
            }

            // Final compliance result
            Console.WriteLine(allCompliant
                ? "The workbook theme is fully compliant with the corporate style guide."
                : "The workbook theme is NOT compliant with the corporate style guide.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
