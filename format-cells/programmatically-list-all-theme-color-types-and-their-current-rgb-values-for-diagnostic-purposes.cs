// Title: Enumerate every Aspose.Cells ThemeColorType and output its RGB components in C#
// AI Prompts: Generate C# code that iterates over the ThemeColorType enum and prints each theme color's R, G, B values using Workbook.GetThemeColor. | Show how to create a dictionary that maps ThemeColorType to System.Drawing.Color for diagnostic inspection in Aspose.Cells. | Provide a method that returns a formatted list of strings like "ThemeColorType: R=..., G=..., B=..." from a newly created Workbook.
// Common Searches: how to get RGB values of all theme colors in Aspose.Cells .NET | C# code to list ThemeColorType enum with corresponding colors | diagnostic script for workbook theme palette using Aspose.Cells | retrieve theme color palette programmatically Aspose.Cells C#
// Tags: GetThemeColor enumeration C# | ThemeColorType RGB extraction Aspose.Cells | workbook theme palette inspection .NET | Aspose.Cells theme color diagnostics | list theme colors programmatically

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For ThemeColorType enumeration

// // Creates a Workbook, loops through every ThemeColorType enum value, obtains each theme color via GetThemeColor, and writes the type together with its R, G, B components to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (no external file needed)
            Workbook workbook = new Workbook();

            // Iterate through all defined theme color types
            foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the theme color for the current type
                Color color = workbook.GetThemeColor(colorType);

                // Output the theme color type and its RGB components
                Console.WriteLine($"{colorType}: R={color.R}, G={color.G}, B={color.B}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
