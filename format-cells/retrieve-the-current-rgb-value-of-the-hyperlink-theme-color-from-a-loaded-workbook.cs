// Title: Retrieve the RGB components of the Hyperlink theme color from an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, calls GetThemeColor for ThemeColorType.Hyperlink, and prints the red, green, and blue values. | Show how to combine the three RGB components of a theme color into a single 0xRRGGBB integer in C# using bitwise operations.
// Common Searches: Aspose.Cells C# get hyperlink theme color RGB values | How to read Excel Hyperlink theme color with Aspose.Cells | Convert Aspose.Cells ThemeColor to hex code in .NET | Extract R G B components from workbook theme color using C# | GetThemeColor Hyperlink example Aspose.Cells
// Tags: Aspose.Cells GetThemeColor Hyperlink RGB | C# extract Excel theme color components | Workbook theme color to hex conversion | Hyperlink theme color retrieval Aspose.Cells | Bitwise combine RGB values C#

using System;
using System.Drawing;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, uses Workbook.GetThemeColor(ThemeColorType.Hyperlink) to obtain the hyperlink theme color, extracts its R, G, B channels, optionally combines them into a 0xRRGGBB integer, and outputs the results.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the Hyperlink theme color
        Color hyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

        // Extract RGB components
        int red   = hyperlinkColor.R;
        int green = hyperlinkColor.G;
        int blue  = hyperlinkColor.B;

        // Combine into a single RGB integer if needed
        int rgb = (red << 16) | (green << 8) | blue;

        // Output the result
        Console.WriteLine($"Hyperlink Theme Color RGB: ({red}, {green}, {blue})");
        Console.WriteLine($"Combined RGB value: 0x{rgb:X6}");
    }
}
