// Title: Import a JSON color scheme and set a custom workbook theme with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load theme colors from a JSON file and apply to workbook | How to use SetThemeColor with JSON-defined colors in Aspose.Cells | Apply custom Excel theme colors from a dictionary in Aspose.Cells .NET | Read JSON color palette and map to ThemeColorType for Excel workbook using Aspose.Cells
// Tags: Aspose.Cells SetThemeColor API | JSON defined Excel theme colors | C# map ThemeColorType enum | custom workbook theme Aspose.Cells | deserialize JSON to dictionary for theme colors

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example creates a new Workbook, reads a JSON file containing theme color names and hex values, deserializes it into a Dictionary<string,string>, converts each hex string to System.Drawing.Color, maps the keys to ThemeColorType enums, applies the colors with Workbook.SetThemeColor, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // ---------- Load JSON-defined color scheme ----------
            string jsonPath = "themeColors.json";

            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            Dictionary<string, string> colorScheme;
            try
            {
                // Read JSON content
                string jsonContent = File.ReadAllText(jsonPath);
                // Deserialize to a dictionary where key is theme color name and value is hex color string
                colorScheme = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonContent);
                if (colorScheme == null)
                {
                    Console.WriteLine("Failed to deserialize color scheme.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading or parsing JSON: {ex.Message}");
                return;
            }

            // ---------- Apply the color scheme to the workbook's theme ----------
            foreach (KeyValuePair<string, string> kvp in colorScheme)
            {
                // Convert the key to ThemeColorType enum (e.g., "Accent1", "Accent2", "Hyperlink", etc.)
                if (Enum.TryParse(kvp.Key, out ThemeColorType themeColor))
                {
                    try
                    {
                        // Convert hex string to System.Drawing.Color
                        Color color = ColorTranslator.FromHtml(kvp.Value);
                        // Set the color in the workbook's theme
                        workbook.SetThemeColor(themeColor, color);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to set color for {kvp.Key}: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Unrecognized theme color type: {kvp.Key}");
                }
            }

            // ---------- Save the workbook ----------
            string outputPath = "WorkbookWithCustomTheme.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
