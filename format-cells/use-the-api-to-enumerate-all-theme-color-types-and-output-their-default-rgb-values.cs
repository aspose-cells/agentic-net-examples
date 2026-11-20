// Title: How to enumerate all ThemeColorType values and display their default RGB colors using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list default theme colors and RGB values | GetThemeColor example for all ThemeColorType entries in .NET | Iterate ThemeColorType enum to retrieve Excel theme palette with Aspose.Cells | How to print Excel theme color RGB values using Aspose.Cells API | Aspose.Cells retrieve workbook theme colors programmatically
// Tags: Aspose.Cells GetThemeColor API | enumerate ThemeColorType enum | default theme color RGB values | C# workbook theme color extraction | list Excel theme colors programmatically

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a Workbook object, loops through every ThemeColorType enum member, calls workbook.GetThemeColor to obtain each theme's default Color, and writes the theme name with its R, G, B components to the console, handling any exceptions that may arise.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (required by Aspose.Cells lifecycle)
            Workbook workbook = new Workbook();

            // Enumerate all ThemeColorType values
            foreach (ThemeColorType themeType in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the default color for the current theme type using the workbook instance
                Color defaultColor = workbook.GetThemeColor(themeType);

                // Output the theme type name and its default RGB values
                Console.WriteLine($"{themeType}: R={defaultColor.R}, G={defaultColor.G}, B={defaultColor.B}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }

        // No saving is needed for this enumeration task
    }
}
