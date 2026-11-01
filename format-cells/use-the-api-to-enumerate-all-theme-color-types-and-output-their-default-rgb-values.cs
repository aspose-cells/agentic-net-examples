// Title: How to enumerate ThemeColorType enum and print default RGB values with Aspose.Cells for .NET
// Common Searches: list all ThemeColorType values with their default RGB using Aspose.Cells C# | get default theme colors from a workbook in Aspose.Cells .NET | retrieve RGB components of theme colors via Workbook.GetThemeColor method | enumerate theme color types in Aspose.Cells and output R G B values | C# code to display default theme palette colors in Aspose.Cells
// Tags: Aspose.Cells ThemeColorType enumeration | Workbook.GetThemeColor usage | default theme color RGB Aspose.Cells | enumerate theme colors .NET | C# Aspose.Cells theme palette extraction

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for ThemeColorType

// The example creates a Workbook, iterates through every ThemeColorType enum member, obtains each type's default RGB color using Workbook.GetThemeColor, and writes the ThemeColorType name together with its R, G, B components to the console, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default theme will be loaded)
            Workbook workbook = new Workbook();

            // Enumerate all values of the ThemeColorType enum
            foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the default color for the current theme color type
                Color rgbColor = workbook.GetThemeColor(colorType);

                // Output the theme color type and its default RGB components
                Console.WriteLine($"{colorType}: R={rgbColor.R}, G={rgbColor.G}, B={rgbColor.B}");
            }
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
