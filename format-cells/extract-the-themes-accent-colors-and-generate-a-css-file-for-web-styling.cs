// Title: Generate a CSS stylesheet from Excel workbook theme accent colors using Aspose.Cells for .NET
// Common Searches: how to read Excel theme accent colors with Aspose.Cells in C# | convert workbook theme colors to hex and create CSS classes using Aspose.Cells | C# code to export Excel theme colors to a CSS file | Aspose.Cells GetThemeColor example for generating CSS stylesheet | extract six accent colors from an Excel file and write a .css file in .NET
// Tags: Aspose.Cells GetThemeColor to hex conversion | generate CSS from Excel theme colors C# | export workbook accent colors as CSS classes | theme accent color extraction Aspose.Cells .NET | create stylesheet from Excel theme using Aspose.Cells | C# write CSS file from Excel theme palette

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Text;

// The example loads an Excel workbook with Aspose.Cells, retrieves the six theme accent colors via GetThemeColor, converts each color to a hexadecimal string, builds .accentN and .accentN-bg CSS class definitions, and writes the resulting stylesheet to a .css file.
class ThemeToCss
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "theme.css";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the six accent colors from the workbook theme
            Color[] accentColors = new Color[6];
            accentColors[0] = workbook.GetThemeColor(ThemeColorType.Accent1);
            accentColors[1] = workbook.GetThemeColor(ThemeColorType.Accent2);
            accentColors[2] = workbook.GetThemeColor(ThemeColorType.Accent3);
            accentColors[3] = workbook.GetThemeColor(ThemeColorType.Accent4);
            accentColors[4] = workbook.GetThemeColor(ThemeColorType.Accent5);
            accentColors[5] = workbook.GetThemeColor(ThemeColorType.Accent6);

            // Build CSS definitions for each accent color
            StringBuilder cssBuilder = new StringBuilder();

            for (int i = 0; i < accentColors.Length; i++)
            {
                Color color = accentColors[i];
                string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

                // Text color class
                cssBuilder.AppendLine($".accent{i + 1} {{ color: {hex}; }}");

                // Background color class
                cssBuilder.AppendLine($".accent{i + 1}-bg {{ background-color: {hex}; }}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the generated CSS to a file
            File.WriteAllText(outputPath, cssBuilder.ToString());

            Console.WriteLine($"CSS file '{outputPath}' has been created.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
