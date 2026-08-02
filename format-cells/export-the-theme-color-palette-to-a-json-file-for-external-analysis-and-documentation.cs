// Title: Export Excel Theme Palette to JSON with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, optionally changes accent colors, reads the first 12 ThemeColorType entries, converts each to an ARGB hex string, serializes the collection to indented JSON, and writes ThemePalette.json while optionally saving the workbook.
// Keywords: Aspose.Cells | C# | Export theme colors | Excel theme palette | JSON serialization | GetThemeColor | SetThemeColor | ThemeColorType | Workbook theme extraction | .NET
// Common Searches: Aspose.Cells export theme colors to JSON | C# read Excel theme palette Aspose | How to get theme colors from workbook using Aspose.Cells | Save Excel theme palette as JSON file | Extract workbook theme colors .NET
// Developer Intent: Generate a JSON file that lists the workbook's theme color palette for external analysis or documentation.
// Use Cases: Document the exact theme colors of an Excel template for a design system. | Compare theme palettes across multiple workbooks by exporting each to JSON. | Supply UI styling tools with the workbook's theme colors to ensure consistent theming. | Audit corporate branding consistency by analyzing exported theme palettes. | Automate creation of style guides from Excel files.
// AI Prompts: Write C# code using Aspose.Cells to read all theme colors from a workbook and output a pretty‑printed JSON file. | Create a method that returns a JSON string containing only theme colors that differ from Aspose.Cells' default palette. | Generate a console application that accepts a workbook path argument and exports its theme palette to a specified JSON file. | Show how to serialize the theme palette dictionary with camelCase property names for JavaScript consumption. | Provide code to batch‑process a folder of Excel files, exporting each workbook's theme palette to separate JSON files.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemePaletteExport
{
    // Creates a workbook, optionally changes accent colors, reads the first 12 ThemeColorType entries, converts each to an ARGB hex string, serializes the collection to indented JSON, and writes ThemePalette.json while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Optional: modify some theme colors to demonstrate export
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 255, 0, 0)); // Red
            workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 0, 255, 0)); // Green
            workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(255, 0, 0, 255)); // Blue

            // Collect all theme colors (there are 12 defined in ThemeColorType enum)
            var themeColors = new Dictionary<string, string>();
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Only consider the first 12 types (0-11) which correspond to theme palette entries
                if ((int)type > 11) break;

                Color color = workbook.GetThemeColor(type);
                // Store as ARGB hex string for readability
                themeColors[type.ToString()] = $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            // Serialize the dictionary to JSON
            string json = JsonSerializer.Serialize(themeColors, new JsonSerializerOptions { WriteIndented = true });

            // Define output file path
            string jsonPath = "ThemePalette.json";

            // Write JSON to file (lifecycle save)
            File.WriteAllText(jsonPath, json);

            // (Optional) Save the workbook to demonstrate that changes persist
            workbook.Save("ThemePaletteDemo.xlsx");

            Console.WriteLine($"Theme palette exported to {jsonPath}");
        }
    }
}
