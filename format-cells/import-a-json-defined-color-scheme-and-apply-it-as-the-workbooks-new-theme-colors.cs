// Title: Apply a JSON‑Defined 12‑Color Palette as a Custom Workbook Theme with Aspose.Cells for .NET
// Description: Read a JSON file containing twelve HTML hex color strings, convert them to System.Drawing.Color objects, create a new Workbook, apply the colors as a custom theme using Aspose.Cells' CustomTheme method, and save the workbook to showcase the new theme.
// Keywords: Aspose.Cells | C# | .NET | JSON color palette | custom workbook theme | Excel theme programmatically | CustomTheme method | import theme from JSON | color scheme Excel | dynamic Excel styling
// Common Searches: Aspose.Cells load custom theme from JSON | C# set Excel workbook theme colors programmatically | apply JSON color palette to Excel using Aspose.Cells | how to use CustomTheme with a JSON file | import 12 hex colors into Aspose.Cells workbook
// Developer Intent: Read a JSON array of twelve hex colors and apply them as a custom theme to a new Aspose.Cells workbook.
// Use Cases: Brand corporate reports by loading brand colors from a JSON configuration and applying them as a workbook theme. | Generate user‑specific themed spreadsheets where the color scheme is supplied in a JSON file. | Validate external theme definitions (exactly 12 colors) before applying them to ensure consistent styling across generated Excel files.
// AI Prompts: Write C# code that reads a JSON array of 12 hex color strings, converts each to System.Drawing.Color, and uses Aspose.Cells.CustomTheme to apply them as a theme named "MyTheme". | Provide error handling for missing JSON files, incorrect array length, or invalid hex strings when creating a custom Excel theme with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeFromJson
{
    // Read a JSON file containing twelve HTML hex color strings, convert them to System.Drawing.Color objects, create a new Workbook, apply the colors as a custom theme using Aspose.Cells' CustomTheme method, and save the workbook to showcase the new theme.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON file that defines the theme colors.
                // The JSON should be an array of 12 color strings (e.g., "#FF0000").
                string jsonPath = "themeColors.json";

                // Verify that the JSON file exists before attempting to read it.
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found: {jsonPath}");
                    return;
                }

                // Read and parse the JSON file.
                string jsonContent = File.ReadAllText(jsonPath);
                string[] colorStrings = JsonSerializer.Deserialize<string[]>(jsonContent);

                if (colorStrings == null || colorStrings.Length != 12)
                {
                    Console.WriteLine("The JSON file must contain exactly 12 color definitions.");
                    return;
                }

                // Convert the string representations to System.Drawing.Color objects.
                Color[] themeColors = new Color[12];
                for (int i = 0; i < 12; i++)
                {
                    // ColorTranslator can handle HTML hex strings.
                    themeColors[i] = ColorTranslator.FromHtml(colorStrings[i]);
                }

                // Create a new workbook (lifecycle rule: create).
                Workbook workbook = new Workbook();

                // Apply the custom theme using the parsed colors (lifecycle rule: modify).
                workbook.CustomTheme("JsonImportedTheme", themeColors);

                // Save the workbook to demonstrate the applied theme (lifecycle rule: save).
                string outputPath = "WorkbookWithJsonTheme.xlsx";
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
