// Title: Create a CSS file with custom properties from Excel theme accent colors using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells read Excel workbook theme accent colors and write them to a .css file | How to export Excel theme palette as CSS custom properties with Aspose.Cells | Generate CSS variables (--accent1, --accent2) from Excel theme using .NET | Retrieve workbook Theme.AccentColors and convert to hex for web styling in C#
// Tags: Aspose.Cells extract theme accent colors | C# generate CSS custom properties from Excel theme | Workbook.Theme.AccentColors to hex conversion | export Excel theme palette as CSS file | System.Drawing.Color hex for web CSS

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, accesses the workbook's Theme object, extracts the AccentColors array, converts each System.Drawing.Color to a hexadecimal string, and writes them as CSS custom properties (--accent1, --accent2, …) into a generated theme-accents.css file.
class ThemeAccentCssGenerator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "theme-accents.css";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the theme using dynamic to avoid compile‑time dependency on Theme type
            dynamic theme = null;
            try
            {
                theme = workbook.Theme;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to retrieve theme: {ex.Message}");
                return;
            }

            if (theme == null)
            {
                Console.WriteLine("The workbook does not contain a theme.");
                return;
            }

            // Get accent colors; cast to Color[] safely
            Color[] accentColors = null;
            try
            {
                accentColors = (Color[])theme.AccentColors;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to obtain accent colors: {ex.Message}");
                return;
            }

            if (accentColors == null || accentColors.Length == 0)
            {
                Console.WriteLine("No accent colors found in the theme.");
                return;
            }

            // Build CSS content using CSS custom properties (variables)
            using (StringWriter cssWriter = new StringWriter())
            {
                cssWriter.WriteLine(":root {");

                for (int i = 0; i < accentColors.Length; i++)
                {
                    // Convert the Color to a hex string like #RRGGBB
                    string hex = ColorTranslator.ToHtml(accentColors[i]);
                    cssWriter.WriteLine($"  --accent{i + 1}: {hex};");
                }

                cssWriter.WriteLine("}");

                // Save the generated CSS to a file
                File.WriteAllText(outputPath, cssWriter.ToString());
            }

            Console.WriteLine($"CSS file '{outputPath}' has been generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
