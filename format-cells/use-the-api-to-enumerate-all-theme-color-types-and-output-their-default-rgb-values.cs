// Title: How to enumerate all ThemeColorType values and retrieve their default RGB colors using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over the ThemeColorType enum and prints each theme color's default RGB components using Workbook.GetThemeColor. | Create a method that returns a Dictionary<ThemeColorType, string> where each value is the hex representation of the default theme color. | Show how to log the names of all theme colors and their RGB values to a text file with Aspose.Cells.
// Common Searches: C# Aspose.Cells retrieve default theme colors for each enum value | list Excel theme colors and their RGB values using Aspose.Cells library | how to display all theme color types with their colors in .NET | sample code to print theme colors from a workbook in Aspose.Cells
// Tags: Aspose.Cells ThemeColorType enumeration | default theme colors with Workbook API | C# extract Excel theme color RGB | list Excel theme color types programmatically | console output of theme colors Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, loops through every ThemeColorType enum member, obtains each type's default color via Workbook.GetThemeColor, and writes the theme name together with its R, G, B values to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (no existing file needed)
                Workbook workbook = new Workbook();

                // Enumerate all ThemeColorType values and output their default RGB values
                foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
                {
                    // Get the default color for the current theme color type
                    Color rgb = workbook.GetThemeColor(colorType);

                    // Print the theme color type and its RGB components
                    Console.WriteLine($"{colorType}: R={rgb.R}, G={rgb.G}, B={rgb.B}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
