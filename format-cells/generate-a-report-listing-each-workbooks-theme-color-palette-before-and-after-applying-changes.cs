// Title: C# utility to list and compare Excel workbook theme palettes before and after modification using Aspose.Cells
// Description: Loads each Excel file with Aspose.Cells, prints the 12‑color theme palette, changes Accent1 and Hyperlink colors, shows the updated palette, and saves a new copy. Handles missing or password‑protected files and supports batch processing.
// Keywords: Aspose.Cells theme colors | C# get Excel theme palette | modify Excel theme colors .NET | theme color report Aspose | batch Excel theme update | retrieve ThemeColorType | set theme color Aspose.Cells | Excel workbook palette C#
// Common Searches: how to read Excel theme colors with Aspose.Cells | C# code to change Excel theme accent color | list workbook theme palette before and after change | batch update Excel theme colors using Aspose | Aspose.Cells get and set ThemeColorType | export Excel theme palette to console C#
// Developer Intent: Create a C# program that reports each workbook’s 12‑color theme palette, applies specific theme color changes, and displays the before‑and‑after values.
// Use Cases: Audit existing workbooks to verify corporate theme colors before a branding rollout. | Confirm that programmatic changes to Accent1 and Hyperlink colors are applied correctly. | Automate bulk updates of theme palettes across multiple Excel files while logging original values.
// AI Prompts: Generate a method that returns a dictionary of ThemeColorType and its ARGB components for a given Workbook. | Provide C# code to write the before‑and‑after theme palette report to a CSV file instead of the console. | Explain how to open password‑protected Excel files with Aspose.Cells and still retrieve their theme colors.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads each Excel file with Aspose.Cells, prints the 12‑color theme palette, changes Accent1 and Hyperlink colors, shows the updated palette, and saves a new copy. Handles missing or password‑protected files and supports batch processing.
class ThemePaletteReport
{
    // Prints the 12 theme colors of a workbook with a label.
    static void PrintThemeColors(Workbook wb, string label)
    {
        Console.WriteLine($"--- {label} Theme Colors ---");
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // Only the first 12 entries correspond to the theme palette.
            if ((int)type > 11) break;

            Color c = wb.GetThemeColor(type);
            Console.WriteLine($"{type}: A={c.A}, R={c.R}, G={c.G}, B={c.B}");
        }
        Console.WriteLine();
    }

    static void Main()
    {
        // List of workbook files to process.
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx"
            // Add more file names as needed.
        };

        foreach (string filePath in workbookFiles)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            Workbook wb = null;
            try
            {
                // Load the workbook. If the file is password‑protected, an exception will be thrown.
                wb = new Workbook(filePath);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Failed to load workbook '{filePath}': {ex.Message}");
                continue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error loading '{filePath}': {ex.Message}");
                continue;
            }

            Console.WriteLine($"Processing workbook: {Path.GetFileName(filePath)}");

            // Report theme colors before changes.
            PrintThemeColors(wb, "Before");

            try
            {
                // Apply a sample change: modify Accent1 and Hyperlink colors.
                wb.SetThemeColor(ThemeColorType.Accent1, Color.Magenta);
                wb.SetThemeColor(ThemeColorType.Hyperlink, Color.Orange);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error modifying theme colors: {ex.Message}");
            }

            // Report theme colors after changes.
            PrintThemeColors(wb, "After");

            try
            {
                // Save the modified workbook.
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(filePath) ?? string.Empty,
                    $"Modified_{Path.GetFileName(filePath)}");
                wb.Save(outputPath);
                Console.WriteLine($"Saved modified workbook to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook '{filePath}': {ex.Message}");
            }

            Console.WriteLine(new string('-', 40));
        }
    }
}
