// Title: Apply a JSON-defined color scheme to an Excel workbook using Aspose.Cells SetThemeColor in C#
// AI Prompts: Read a JSON file that maps ThemeColorType names to hex strings, convert each hex value to System.Drawing.Color, and assign it to the workbook with Workbook.SetThemeColor. | Refactor the code to accept a JSON string or stream instead of a file path, then apply the parsed theme colors to an existing Workbook instance. | Enhance the example with validation that skips entries with invalid hex formats or unknown ThemeColorType keys while logging detailed warnings.
// Common Searches: C# Aspose.Cells load custom theme colors from JSON file | How to set Excel theme colors programmatically with Aspose.Cells | Parse hex color values from JSON and apply to workbook theme in .NET | Aspose.Cells SetThemeColor using dictionary of color names
// Tags: setthemecolor from json Aspose.Cells | json-defined excel theme colors c# | convert hex to System.Drawing.Color c# | apply custom workbook theme Aspose.Cells | themecolortype mapping json c#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;

// The sample creates a new Workbook, reads a JSON file containing ThemeColorType keys and hex color values, converts each hex string to a System.Drawing.Color, maps the keys to ThemeColorType, applies the colors via Workbook.SetThemeColor, and saves the workbook, with error handling for missing files, invalid colors, and unrecognized keys.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook.
            Workbook workbook = new Workbook();

            // Path to the JSON file that defines theme colors.
            string jsonPath = "themeColors.json";

            // Verify that the JSON file exists before attempting to read it.
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            // Read and deserialize the JSON content into a dictionary.
            string json = File.ReadAllText(jsonPath);
            Dictionary<string, string> colorMap = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

            if (colorMap == null)
            {
                Console.WriteLine("Failed to parse JSON or JSON is empty.");
                return;
            }

            // Apply each color to the workbook's theme.
            foreach (KeyValuePair<string, string> entry in colorMap)
            {
                // Convert the key (e.g., "Accent1") to the ThemeColorType enum.
                if (Enum.TryParse(entry.Key, out ThemeColorType themeColor))
                {
                    try
                    {
                        // Convert the hex string (e.g., "#FF0000") to a System.Drawing.Color.
                        Color color = ColorTranslator.FromHtml(entry.Value);
                        // Set the theme color directly on the workbook.
                        workbook.SetThemeColor(themeColor, color);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing color value '{entry.Value}' for key '{entry.Key}': {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Unrecognized theme color key: {entry.Key}");
                }
            }

            // Save the workbook with the new theme applied.
            string outputPath = "WorkbookWithCustomTheme.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
