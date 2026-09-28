// Title: Create a CSS file with Excel theme accent colors using Aspose.Cells in C#
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, reads the six theme accent colors via Workbook.GetThemeColor, converts them to hex, and writes them as CSS custom properties in a .css file. | Add error handling to verify the input workbook exists and output a clear message before attempting to generate the CSS stylesheet. | Encapsulate the accent‑color extraction into a reusable method that returns a dictionary mapping "--accent1" … "--accent6" to their hex values.
// Common Searches: how to read Excel theme accent colors with Aspose.Cells in C# | C# generate CSS variables from workbook theme colors | Aspose.Cells GetThemeColor example for creating a stylesheet | export Excel theme colors to a .css file using .NET | convert Excel theme colors to hex for web styling in C#
// Tags: Aspose.Cells GetThemeColor C# | export workbook theme colors to CSS | generate CSS custom properties from workbook theme | C# extract workbook theme colors for web styling | write CSS root variables using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook, retrieves its six theme accent colors via Workbook.GetThemeColor, converts each to a hex string, and writes them as CSS custom properties inside a :root selector in a generated .css file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "theme-colors.css";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Build CSS content with theme accent colors
            StringBuilder cssBuilder = new StringBuilder();
            cssBuilder.AppendLine(":root {");

            // Aspose.Cells provides up to 6 accent colors in a theme.
            // Use Workbook.GetThemeColor to retrieve each accent color.
            for (int i = 0; i < 6; i++)
            {
                ThemeColorType accentType = (ThemeColorType)Enum.Parse(typeof(ThemeColorType), $"Accent{i + 1}");
                Color accentColor = workbook.GetThemeColor(accentType);
                string hex = ColorTranslator.ToHtml(accentColor);
                cssBuilder.AppendLine($"  --accent{i + 1}: {hex};");
            }

            cssBuilder.AppendLine("}");

            // Write the CSS content to a file
            File.WriteAllText(outputPath, cssBuilder.ToString());

            Console.WriteLine($"CSS file '{outputPath}' has been generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
