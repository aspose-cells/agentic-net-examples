// Title: C# code to enumerate all ThemeColorType values and display their RGB components with Aspose.Cells
// Common Searches: C# enumerate ThemeColorType enum and get RGB values using Aspose.Cells | how to retrieve workbook theme colors programmatically in Aspose.Cells .NET | list theme palette colors and RGB in Aspose.Cells for .NET | Aspose.Cells GetThemeColor example for all theme colors | diagnostic output of theme color RGB values in a new workbook Aspose.Cells
// Tags: Aspose.Cells GetThemeColor enumeration C# | list ThemeColorType RGB values Aspose.Cells | retrieve workbook theme palette programmatically | diagnostic theme color extraction .NET | enumerate theme colors using Aspose.Cells API

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a new Workbook, loops through every ThemeColorType enum member, calls Workbook.GetThemeColor to obtain the actual Color, and prints each type with its R, G, B values to the console, with error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (no need to load an existing file)
            Workbook workbook = new Workbook();

            // Iterate through all defined ThemeColorType values and retrieve their colors
            foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Get the actual color for the current theme color type
                Color color = workbook.GetThemeColor(colorType);

                // Output the theme color type and its RGB components
                Console.WriteLine($"{colorType}: R={color.R}, G={color.G}, B={color.B}");
            }
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
