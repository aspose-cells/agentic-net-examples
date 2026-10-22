// Title: Change the Dark2 theme color in an Excel workbook using a JSON configuration file with Aspose.Cells for .NET
// Common Searches: how to read a JSON file and set the Dark2 theme color in an Aspose.Cells workbook C# | Aspose.Cells C# update default style foreground and background color from hex string | apply external theme configuration to an Excel file using Aspose.Cells .NET | change Excel theme colors programmatically with Aspose.Cells and JSON | load themeConfig.json and modify workbook theme colors in C# Aspose.Cells
// Tags: update Dark2 theme color Aspose.Cells | read JSON theme configuration C# | set workbook default style color Aspose.Cells | apply external theme to Excel file .NET | modify Excel theme colors programmatically

using System;
using System.IO;
using System.Text.Json;
using System.Drawing;
using Aspose.Cells;

// The example loads an existing Excel workbook, reads a JSON file containing a hex value for the Dark2 theme color, converts the hex string to a System.Drawing.Color, assigns that color to the workbook's default style foreground and background, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string configPath = "themeConfig.json";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the configuration file exists
            if (!File.Exists(configPath))
                throw new FileNotFoundException($"Theme configuration file not found: {configPath}");

            // Read and deserialize the theme configuration
            string configJson = File.ReadAllText(configPath);
            ThemeConfig? config = JsonSerializer.Deserialize<ThemeConfig>(configJson);
            if (config?.Dark2 == null)
                throw new InvalidOperationException("Dark2 color is missing in the configuration.");

            // Convert the hex color string to System.Drawing.Color
            Color newDark2 = ColorTranslator.FromHtml(config.Dark2);

            // Apply the new color to the workbook's default style
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.ForegroundColor = newDark2;
            defaultStyle.BackgroundColor = newDark2;
            workbook.DefaultStyle = defaultStyle; // ensure changes are saved

            // NOTE: Aspose.Cells does not expose a direct Styles collection in some versions.
            // If needed, individual styles can be accessed via workbook.GetStyle or by iterating worksheets.
            // For simplicity, only the default style is updated here.

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple class to map the JSON configuration
    private class ThemeConfig
    {
        public string? Dark2 { get; set; }
    }
}
