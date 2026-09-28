// Title: Retrieve Excel theme accent colors and convert them to HSL using Aspose.Cells in C#
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, reads ThemeColorType.Accent1‑Accent6 via GetThemeColor, converts each System.Drawing.Color to HSL, and prints the results. | Add file‑existence validation and comprehensive exception handling to a console application that extracts theme accent colors from an Excel file using Aspose.Cells. | Implement a reusable static method that transforms a System.Drawing.Color (RGB) into HSL and apply it to all theme accent colors obtained from a workbook.
// Common Searches: Aspose.Cells C# how to read Excel theme accent colors | convert System.Drawing.Color to HSL in .NET | C# example for Workbook.GetThemeColor returning RGB values | log Excel theme palette HSL values using Aspose.Cells | handle missing Excel file when extracting theme colors with Aspose.Cells
// Tags: Aspose.Cells GetThemeColor accent extraction | C# RGB to HSL conversion method | Excel theme palette reading .NET | console logging HSL values | file existence check Aspose.Cells workbook

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// This C# console sample loads an .xlsx workbook with Aspose.Cells, retrieves the six theme accent colors via Workbook.GetThemeColor, converts each RGB color to HSL using a custom conversion method, and writes both RGB and HSL values to the console while handling missing files and other exceptions.
class ThemeAccentExtractor
{
    // Convert a System.Drawing.Color (RGB) to HSL.
    // Returns a tuple: (Hue [0-360], Saturation [0-1], Lightness [0-1])
    static (float H, float S, float L) RgbToHsl(Color rgb)
    {
        float r = rgb.R / 255f;
        float g = rgb.G / 255f;
        float b = rgb.B / 255f;

        float max = Math.Max(r, Math.Max(g, b));
        float min = Math.Min(r, Math.Min(g, b));
        float delta = max - min;

        // Lightness
        float l = (max + min) / 2f;

        float h = 0f;
        float s = 0f;

        if (delta != 0f)
        {
            // Saturation
            s = l < 0.5f ? delta / (max + min) : delta / (2f - max - min);

            // Hue
            if (max == r)
                h = ((g - b) / delta) % 6f;
            else if (max == g)
                h = ((b - r) / delta) + 2f;
            else // max == b
                h = ((r - g) / delta) + 4f;

            h *= 60f; // convert to degrees
            if (h < 0f)
                h += 360f;
        }

        return (h, s, l);
    }

    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the six accent theme colors
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Log each accent color and its HSL representation
            for (int i = 0; i < accentTypes.Length; i++)
            {
                Color rgb = workbook.GetThemeColor(accentTypes[i]);
                var (h, s, l) = RgbToHsl(rgb);
                Console.WriteLine($"Accent {i + 1}: RGB({rgb.R}, {rgb.G}, {rgb.B}) => HSL({h:F1}°, {s:P1}, {l:P1})");
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
