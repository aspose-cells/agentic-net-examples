using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeImport
{
    // Model representing the JSON color scheme
    public class ThemeScheme
    {
        public string[] colors { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the JSON file containing 12 color definitions (hex strings)
                string jsonPath = "themeColors.json";

                // Verify that the JSON file exists
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"File not found: {jsonPath}");
                    return;
                }

                // Read and deserialize the JSON file
                ThemeScheme scheme = JsonSerializer.Deserialize<ThemeScheme>(File.ReadAllText(jsonPath));

                // Validate that exactly 12 colors are provided
                if (scheme?.colors == null || scheme.colors.Length != 12)
                {
                    Console.WriteLine("The JSON must contain exactly 12 color values.");
                    return;
                }

                // Convert hex strings to System.Drawing.Color objects
                Color[] themeColors = new Color[12];
                for (int i = 0; i < 12; i++)
                {
                    // Supports formats like "#RRGGBB" or "RRGGBB"
                    string hex = scheme.colors[i].TrimStart('#');
                    if (hex.Length != 6)
                    {
                        Console.WriteLine($"Invalid color format at index {i}: {scheme.colors[i]}");
                        return;
                    }

                    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                    int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                    themeColors[i] = Color.FromArgb(255, r, g, b); // Fully opaque
                }

                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Apply the custom theme using the imported colors (lifecycle rule: modify)
                workbook.CustomTheme("ImportedJsonTheme", themeColors);

                // Optional: demonstrate the theme by applying an accent color to a cell
                Worksheet sheet = workbook.Worksheets[0];
                Cell demoCell = sheet.Cells["A1"];
                demoCell.PutValue("Theme Applied");
                Style style = workbook.CreateStyle();
                style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                demoCell.SetStyle(style);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "WorkbookWithImportedTheme.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with the imported theme: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}