// Title: How to export an Excel workbook's theme color palette to a JSON file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export workbook theme colors to JSON | Retrieve Excel theme palette using Aspose.Cells and save as JSON file | How to get theme color hex values from an .xlsx with Aspose.Cells in .NET | Aspose.Cells GetThemeColor fallback to Theme.ColorScheme example | Serialize Excel theme color scheme to JSON using C#
// Tags: Aspose.Cells theme palette extraction | Workbook GetThemeColor API | Theme.ColorScheme GetColor fallback | Excel theme colors hex conversion | JSON output of Excel theme palette

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.Json;

// The example loads an .xlsx workbook with Aspose.Cells, extracts the built‑in theme colors (Dark1, Light1, Dark2, Light2, Accent1‑Accent6, Hyperlink, FollowedHyperlink) using either Workbook.GetThemeColor or the older Theme.ColorScheme.GetColor via reflection, converts each Color to a hexadecimal string, stores them in a dictionary, serializes the dictionary to indented JSON, and writes the result to "themePalette.json".
class ExportThemePalette
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "themePalette.json";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Resolve the ThemeColorType enum via reflection (covers different Aspose.Cells versions)
            Type themeColorEnum = Type.GetType("Aspose.Cells.Drawing.ThemeColorType, Aspose.Cells");
            if (themeColorEnum == null)
            {
                Console.WriteLine("Error: ThemeColorType enum not found in the loaded Aspose.Cells assembly.");
                return;
            }

            // List of theme color names to export
            string[] themeColorNames = new string[]
            {
                "Dark1", "Light1", "Dark2", "Light2",
                "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
                "Hyperlink", "FollowedHyperlink"
            };

            var palette = new Dictionary<string, string>();

            // Try to locate a direct GetThemeColor method on Workbook (newer API)
            MethodInfo getThemeColorMethod = workbook.GetType().GetMethod("GetThemeColor", new[] { themeColorEnum });

            // If not available, prepare to use Theme -> ColorScheme -> GetColor (older API)
            PropertyInfo themeProperty = workbook.GetType().GetProperty("Theme");
            object themeInstance = themeProperty?.GetValue(workbook);
            PropertyInfo colorSchemeProperty = themeInstance?.GetType().GetProperty("ColorScheme");
            object colorSchemeInstance = colorSchemeProperty?.GetValue(themeInstance);
            MethodInfo getColorMethod = colorSchemeInstance?.GetType().GetMethod("GetColor", new[] { themeColorEnum });

            foreach (string name in themeColorNames)
            {
                // Parse enum value from name
                object enumValue = Enum.Parse(themeColorEnum, name);

                Color color = Color.Empty;

                // Prefer direct Workbook method if present
                if (getThemeColorMethod != null)
                {
                    object result = getThemeColorMethod.Invoke(workbook, new[] { enumValue });
                    if (result is Color c) color = c;
                }
                // Fallback to Theme -> ColorScheme method
                else if (getColorMethod != null)
                {
                    object result = getColorMethod.Invoke(colorSchemeInstance, new[] { enumValue });
                    if (result is Color c) color = c;
                }

                // If a valid color was retrieved, add it to the palette
                if (color != Color.Empty)
                {
                    string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                    palette[name] = hex;
                }
                else
                {
                    Console.WriteLine($"Warning: Unable to retrieve color for \"{name}\".");
                }
            }

            // Serialize the palette dictionary to formatted JSON
            string json = JsonSerializer.Serialize(palette, new JsonSerializerOptions { WriteIndented = true });

            // Save the JSON to a file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Theme palette exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
