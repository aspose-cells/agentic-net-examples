// Title: C# example to load an XLSX workbook with Aspose.Cells and list all theme color definitions
// AI Prompts: Write C# code that opens a .xlsx file using Aspose.Cells, iterates through the ThemeColorType enum, and prints each theme color's ARGB components. | Create a C# routine that extracts the workbook's theme colors via Workbook.GetThemeColor and writes the definitions to a JSON file. | Enhance the theme‑color extraction to continue enumeration when a ThemeColorType cannot be retrieved, logging a warning instead of throwing.
// Common Searches: how to retrieve Excel theme colors with Aspose.Cells in C# | list all ThemeColorType values and their ARGB values using Aspose.Cells .NET | extract theme palette from an XLSX file programmatically with Aspose.Cells | C# code sample for reading workbook theme colors via GetThemeColor
// Tags: Aspose.Cells GetThemeColor C# | extract Excel theme palette .NET | enumerate ThemeColorType values Aspose.Cells | load XLSX workbook theme colors Aspose | theme color definitions ARGB Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The sample loads an XLSX workbook with Aspose.Cells, iterates over every ThemeColorType enum value, obtains each color using Workbook.GetThemeColor, formats the ARGB components, and outputs the complete list of theme color definitions.
class ThemeExtractor
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect theme color definitions using Workbook.GetThemeColor
            List<string> themeColors = new List<string>();

            foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
            {
                try
                {
                    Color color = workbook.GetThemeColor(colorType);
                    string definition = $"{colorType}: ARGB({color.A}, {color.R}, {color.G}, {color.B})";
                    themeColors.Add(definition);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to get color for {colorType}: {ex.Message}");
                }
            }

            // Output the extracted theme color definitions
            Console.WriteLine("Theme color definitions:");
            foreach (string def in themeColors)
            {
                Console.WriteLine(def);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
