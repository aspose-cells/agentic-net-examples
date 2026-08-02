// Title: C# – Update Excel Dark2 Theme Color from JSON and Refresh Dependent Styles with Aspose.Cells
// Description: Read a hex value from a JSON file, set the workbook's Dark2 theme color (fallback to Accent1 if missing), reapply all styles that reference the theme, and save the updated Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells set theme color C# | update Dark2 theme Excel | refresh theme dependent styles | load theme color from JSON | fallback to Accent1 theme | Excel theme customization .NET | C# Excel theming example | Aspose.Cells workbook styling
// Common Searches: change Dark2 theme color Aspose.Cells C# | apply JSON color to Excel theme using Aspose | refresh cell styles after theme update Aspose.Cells | fallback to Accent1 when Dark2 not present | C# example for updating Excel theme colors
// Developer Intent: Modify a workbook’s Dark2 theme color based on a JSON configuration and ensure every style that uses that theme color is refreshed.
// Use Cases: Apply a corporate brand color to the Dark2 theme across an existing report workbook. | Allow end‑users to define theme colors in a config file without manually editing each cell style. | Maintain visual consistency by falling back to Accent1 when a workbook lacks a Dark2 theme.
// AI Prompts: Generate C# code that reads a hex color from a JSON file, sets the Dark2 theme in an Aspose.Cells workbook, and refreshes all dependent styles. | Show how to update multiple theme colors from a JSON configuration using the Workbook.ThemeColors collection. | Explain how to detect the absence of the Dark2 theme and programmatically fallback to Accent1 when applying a new theme color with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Read a hex value from a JSON file, set the workbook's Dark2 theme color (fallback to Accent1 if missing), reapply all styles that reference the theme, and save the updated Excel file using Aspose.Cells for .NET.
class UpdateDark2Theme
{
    static void Main()
    {
        try
        {
            // Load configuration file that contains the new Dark2 color in hex format.
            const string configPath = "themeConfig.json";
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Configuration file not found: {configPath}");
                return;
            }

            string json = File.ReadAllText(configPath);
            ThemeConfig? config = JsonSerializer.Deserialize<ThemeConfig>(json);
            if (config == null || string.IsNullOrWhiteSpace(config.Dark2Hex))
            {
                Console.WriteLine("Invalid configuration. Ensure Dark2Hex is present.");
                return;
            }

            // Convert the hex string to a System.Drawing.Color.
            Color dark2Color = ColorTranslator.FromHtml(config.Dark2Hex);

            // Verify the input workbook exists before loading.
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Load the existing workbook (lifecycle rule: load).
            Workbook workbook = new Workbook(inputPath);

            // Determine which theme color to update.
            // Prefer Dark2; if not available, fall back to Accent1.
            ThemeColorType targetTheme = Enum.TryParse<ThemeColorType>("Dark2", out var parsedDark2)
                ? parsedDark2
                : ThemeColorType.Accent1;

            // Update the selected theme color.
            workbook.SetThemeColor(targetTheme, dark2Color);

            // Refresh all styles that depend on the target theme color.
            RefreshThemeDependentStyles(workbook, targetTheme);

            // Save the modified workbook (lifecycle rule: save).
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved with updated {targetTheme} theme color to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Reapplies styles that reference the specified theme color to ensure they reflect the change.
    private static void RefreshThemeDependentStyles(Workbook workbook, ThemeColorType targetType)
    {
        try
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        Cell cell = cells[r, c];
                        if (cell == null) continue;

                        Style style = cell.GetStyle();
                        bool needsRefresh = false;

                        // Font theme color
                        ThemeColor? fontTheme = style.Font.ThemeColor;
                        if (fontTheme != null && fontTheme.ColorType == targetType)
                            needsRefresh = true;

                        // Background theme color
                        ThemeColor? bgTheme = style.BackgroundThemeColor;
                        if (bgTheme != null && bgTheme.ColorType == targetType)
                            needsRefresh = true;

                        // Foreground theme color
                        ThemeColor? fgTheme = style.ForegroundThemeColor;
                        if (fgTheme != null && fgTheme.ColorType == targetType)
                            needsRefresh = true;

                        // Reapply the style if it uses the target theme color.
                        if (needsRefresh)
                            cell.SetStyle(style);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while refreshing styles: {ex.Message}");
        }
    }

    // Simple POCO to map the JSON configuration.
    private class ThemeConfig
    {
        public string? Dark2Hex { get; set; }
    }
}
