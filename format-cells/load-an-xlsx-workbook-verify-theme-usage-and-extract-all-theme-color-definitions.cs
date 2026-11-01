// Title: Extract all theme color definitions from an XLSX workbook using Aspose.Cells for .NET
// Common Searches: how to read Excel theme colors with Aspose.Cells in C# | C# code to list workbook theme color values using GetThemeColor | retrieve all ThemeColorType entries from an XLSX file with Aspose.Cells | extract theme palette colors from a .xlsx workbook programmatically | Aspose.Cells get theme color hex codes from workbook
// Tags: Aspose.Cells GetThemeColor usage | enumerate ThemeColorType in .NET | extract Excel theme palette colors C# | load XLSX workbook with Aspose.Cells | convert theme colors to hex representation

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

// The example verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, iterates over every ThemeColorType enum value, retrieves each theme color via Workbook.GetThemeColor, stores the results in a dictionary, and prints each color name with its hexadecimal RGB value.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file '{inputFile}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the XLSX workbook
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Extract all theme color definitions using Workbook.GetThemeColor
        var themeColors = new Dictionary<string, Color>();

        foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
        {
            try
            {
                // Retrieve the color for the current theme color type
                Color color = workbook.GetThemeColor(colorType);
                themeColors.Add(colorType.ToString(), color);
            }
            catch (Exception ex)
            {
                // If a specific color type is not available, skip it
                Console.WriteLine($"Unable to retrieve color for {colorType}: {ex.Message}");
            }
        }

        // Output the extracted theme colors
        foreach (var kvp in themeColors)
        {
            // Display color name and its hex representation
            Console.WriteLine($"{kvp.Key}: #{kvp.Value.R:X2}{kvp.Value.G:X2}{kvp.Value.B:X2}");
        }
    }
}
