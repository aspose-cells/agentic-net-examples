// Title: C# – Export Aspose.Cells Theme Color Palette to JSON
// Description: Creates a new Workbook, reads its 56‑color theme palette via Workbook.Colors, builds a JSON array containing the index, ARGB hex string and individual RGBA components, and writes the result to ThemePalette.json. The example also demonstrates optional workbook saving.
// Keywords: Aspose.Cells | C# | theme palette export | Workbook.Colors | JSON serialization | System.Text.Json | Excel theme colors | color palette extraction | .NET
// Common Searches: Aspose.Cells export theme colors C# | Get workbook theme palette as JSON | Serialize Excel theme palette Aspose .NET | How to extract 56 theme colors from Aspose.Cells | Save Aspose.Cells theme palette to file
// Developer Intent: Export the workbook's 56‑entry theme color palette to a JSON file for external analysis or documentation.
// Use Cases: Document all theme colors used in generated Excel files. | Compare palettes across multiple workbooks to enforce visual consistency. | Feed the exported colors into a design system or UI style guide. | Automate testing of Excel color schemes in CI pipelines.
// AI Prompts: Write C# code that reads Aspose.Cells Workbook.Colors and outputs a formatted JSON file. | Provide a reusable method that returns the theme palette as a JSON string and optionally saves it to a given path. | Explain how to add custom color names or filter specific indices when exporting the palette.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemePaletteExport
{
    // Creates a new Workbook, reads its 56‑color theme palette via Workbook.Colors, builds a JSON array containing the index, ARGB hex string and individual RGBA components, and writes the result to ThemePalette.json. The example also demonstrates optional workbook saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the mandated create rule)
            Workbook workbook = new Workbook();

            // Retrieve the 56‑entry theme color palette
            Color[] palette = workbook.Colors;

            // Prepare a simple DTO for JSON serialization
            var paletteData = new List<object>();
            for (int i = 0; i < palette.Length; i++)
            {
                Color c = palette[i];
                paletteData.Add(new
                {
                    Index = i,
                    // Store color as ARGB hex string for readability
                    ARGB = $"#{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}",
                    // Also expose individual components if needed
                    A = c.A,
                    R = c.R,
                    G = c.G,
                    B = c.B
                });
            }

            // Serialize the palette to a formatted JSON string
            string json = JsonSerializer.Serialize(paletteData, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON to a file
            string outputPath = "ThemePalette.json";
            File.WriteAllText(outputPath, json);

            // Optionally, save the workbook (demonstrates the mandated save rule)
            workbook.Save("ThemePaletteDemo.xlsx");

            Console.WriteLine($"Theme palette exported to '{outputPath}'.");
        }
    }
}
