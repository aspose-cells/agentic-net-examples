// Title: Update the Dark2 theme color from a JSON configuration and apply it to every cell in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load a JSON file that defines a Dark2 hex value, convert it to a System.Drawing.Color, assign this color to the workbook's theme, and iterate through all worksheets to set each cell's foreground color with Aspose.Cells. | Write a C# helper that accepts a configuration path, updates the workbook's Dark2 theme color, and refreshes all cell styles across the file while handling missing files and invalid color data.
// Common Searches: how to change the Dark2 theme color in an Excel workbook with Aspose.Cells C# | read hex color from json and apply to all cells using Aspose.Cells .NET | programmatically update Excel theme colors and refresh styles in .NET | apply configuration driven theme changes to multiple worksheets with Aspose.Cells
// Tags: Aspose.Cells modify theme colors | C# parse hex color from JSON file | apply foreground color to all cells in worksheet | refresh cell styles after theme change | Excel workbook theme update .NET

using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The C# program loads an existing Excel workbook, reads a JSON configuration to obtain a hex value for the Dark2 theme color, converts it to a System.Drawing.Color, and then iterates through every cell in each worksheet to set the ForegroundColor to the new value. It saves the modified workbook as output.xlsx and includes error handling for missing files and invalid configuration data.
class Program
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify configuration file exists
            const string configPath = "config.json";
            if (!File.Exists(configPath))
                throw new FileNotFoundException($"Configuration file not found: {configPath}");

            // Read and deserialize configuration
            string configContent = File.ReadAllText(configPath);
            Config? config = JsonSerializer.Deserialize<Config>(configContent);
            if (config == null || string.IsNullOrWhiteSpace(config.Dark2))
                throw new InvalidDataException("Invalid configuration: Dark2 color is missing.");

            // Convert the hex color string to a System.Drawing.Color
            Color newDark2Color = ColorTranslator.FromHtml(config.Dark2);

            // Apply the new color to all cells (fallback when Theme API is unavailable)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Style style = cells[row, col].GetStyle();
                        // Set foreground color to the new Dark2 color
                        style.ForegroundColor = newDark2Color;
                        cells[row, col].SetStyle(style);
                    }
                }
            }

            // Save the modified workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple configuration class matching the JSON structure
    private class Config
    {
        public string? Dark2 { get; set; }
    }
}
