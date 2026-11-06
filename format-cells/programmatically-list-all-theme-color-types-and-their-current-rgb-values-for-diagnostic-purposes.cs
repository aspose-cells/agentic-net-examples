// Title: How to enumerate all ThemeColorType values and display their RGB colors with Aspose.Cells for .NET
// Common Searches: C# code to get RGB values of each ThemeColorType using Aspose.Cells | list workbook theme palette colors Aspose.Cells .NET | retrieve theme color enumeration and RGB components in Aspose.Cells | diagnostic output of theme colors from a new workbook with Aspose.Cells | Aspose.Cells GetThemeColor example for all theme colors
// Tags: Aspose.Cells GetThemeColor enumeration | C# extract ThemeColorType RGB values | diagnostic theme palette inspection .NET | Aspose.Cells workbook theme color extraction | theme color type retrieval Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new Workbook, iterates over every ThemeColorType enum value, obtains each theme color via GetThemeColor, formats its R, G, B components, and writes the type and RGB values to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (no need to load an existing file)
            Workbook workbook = new Workbook();

            // Iterate through all ThemeColorType enum values
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the color associated with the current theme type
                Color color = workbook.GetThemeColor(type);

                // Format RGB components for display
                string rgb = $"R:{color.R}, G:{color.G}, B:{color.B}";

                // Output the theme type and its RGB value
                Console.WriteLine($"{type}: {rgb}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
