// Title: C# – Load a JSON Color Palette and Set a Custom Workbook Theme with Aspose.Cells
// Description: The sample reads a JSON file containing twelve hex color strings, converts each entry to a System.Drawing.Color, builds a theme via Workbook.CustomTheme, applies an accent style to a cell, and writes the result to an XLSX workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# JSON theme | CustomTheme method | Excel color palette programmatically | load hex colors from JSON | Workbook.CustomTheme example | Aspose.Cells .NET tutorial | JSON defined colors for Excel | apply theme colors in code | Excel workbook styling
// Common Searches: how to create an Excel theme from a JSON file using Aspose.Cells | Aspose.Cells C# read hex colors array and set workbook theme | example of Workbook.CustomTheme with colors loaded from JSON | apply custom color scheme to Excel workbook programmatically | C# load JSON palette and use it in Aspose.Cells
// Developer Intent: Read a JSON array of twelve hex values and use it to build and apply a custom theme to an Aspose.Cells workbook.
// Use Cases: Implement brand‑consistent reports by importing a JSON‑based color palette at runtime. | Enable multi‑tenant SaaS platforms to generate workbooks with tenant‑specific themes stored in JSON files. | Allow end‑users to switch visual styles by selecting different JSON color schemes without recompiling code.
// AI Prompts: Generate C# code that parses a JSON array of 12 hex colors and creates a custom Excel theme with Aspose.Cells. | Explain validation steps for a JSON color list before calling Workbook.CustomTheme in a .NET application. | Show how to apply an Accent1 style from a JSON‑derived theme to a specific cell using Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeFromJson
{
    // The sample reads a JSON file containing twelve hex color strings, converts each entry to a System.Drawing.Color, builds a theme via Workbook.CustomTheme, applies an accent style to a cell, and writes the result to an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON file that defines the theme colors.
                // Expected format: an array of 12 color strings in hex notation, e.g. ["#FF0000", "#00FF00", ...]
                string jsonPath = "themeColors.json";

                // Verify that the JSON file exists to avoid FileNotFoundException.
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found: {jsonPath}");
                    return;
                }

                // Read the JSON content.
                string jsonContent = File.ReadAllText(jsonPath);

                // Deserialize the JSON array into a string[].
                string[] hexColors = JsonSerializer.Deserialize<string[]>(jsonContent);

                // Validate that exactly 12 colors are provided.
                if (hexColors == null || hexColors.Length != 12)
                {
                    Console.WriteLine("The JSON file must contain exactly 12 color definitions.");
                    return;
                }

                // Convert hex strings to System.Drawing.Color objects.
                Color[] themeColors = new Color[12];
                for (int i = 0; i < 12; i++)
                {
                    // ColorTranslator can parse HTML hex color strings.
                    themeColors[i] = ColorTranslator.FromHtml(hexColors[i]);
                }

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Apply the custom theme using the colors from JSON.
                workbook.CustomTheme("JsonImportedTheme", themeColors);

                // Demonstrate the theme by applying an accent color to a cell.
                Worksheet sheet = workbook.Worksheets[0];
                Cell demoCell = sheet.Cells["A1"];
                demoCell.PutValue("Theme Applied from JSON");

                Style style = workbook.CreateStyle();
                // Use Accent1 from the newly defined theme.
                style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                demoCell.SetStyle(style);

                // Save the workbook with the new theme.
                string outputPath = "CustomThemeWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
