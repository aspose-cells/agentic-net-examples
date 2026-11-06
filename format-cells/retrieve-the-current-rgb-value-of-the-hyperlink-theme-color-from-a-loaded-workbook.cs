// Title: How to read the Hyperlink theme color RGB value from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get hyperlink theme color RGB from workbook | retrieve Excel theme hyperlink color using Aspose.Cells .NET API | Aspose.Cells GetThemeColor Hyperlink example in C# | read workbook theme colors as RGB values with Aspose.Cells | extract hyperlink theme color hex code from Excel file using Aspose.Cells
// Tags: Aspose.Cells GetThemeColor Hyperlink | Excel workbook theme color extraction C# | hyperlink theme RGB retrieval Aspose.Cells | C# read Excel theme colors Aspose | convert theme color to hex Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook with Aspose.Cells, calls GetThemeColor(ThemeColorType.Hyperlink) to obtain the hyperlink theme color, extracts its R, G, B components, combines them into a 0xRRGGBB value, and outputs both the individual RGB values and the hexadecimal representation.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the Hyperlink theme color from the workbook's theme
            Color hyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

            // Retrieve individual RGB components
            int red = hyperlinkColor.R;
            int green = hyperlinkColor.G;
            int blue = hyperlinkColor.B;

            // Combine into a single RGB integer (0xRRGGBB)
            int rgbValue = (red << 16) | (green << 8) | blue;

            // Output the RGB values
            Console.WriteLine($"Hyperlink Theme Color - R:{red} G:{green} B:{blue}");
            Console.WriteLine($"RGB Hex: #{rgbValue:X6}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
