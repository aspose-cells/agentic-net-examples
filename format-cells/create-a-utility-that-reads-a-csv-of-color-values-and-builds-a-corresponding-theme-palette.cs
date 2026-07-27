// Title: Build a Custom Excel Theme from a CSV Palette with Aspose.Cells for .NET (C#)
// Description: A C# utility that reads a CSV file containing exactly 12 colors (hex "#RRGGBB" or "R,G,B"), validates the entries, creates a new Workbook, applies the colors as a custom theme via Aspose.Cells, and saves the themed workbook to a specified location.
// Keywords: Aspose.Cells | C# custom Excel theme | CSV color palette | Excel theme programmatically | Aspose.Cells CustomTheme | read hex RGB CSV | generate workbook theme .NET
// Common Searches: how to create an Excel theme from CSV using Aspose.Cells | C# read hex and RGB colors from a file and apply as workbook theme | Aspose.Cells custom theme requires 12 colors | generate Excel theme programmatically in .NET | load color palette CSV into Aspose.Cells workbook
// Developer Intent: Create an Excel workbook whose theme is defined by a 12‑color palette loaded from a CSV file using Aspose.Cells.
// Use Cases: Transform a brand‑color CSV into a reusable Excel theme for all automated reports. | Batch‑process spreadsheets, applying a consistent palette defined in configuration files. | Validate user‑provided color lists and enforce a standard look across workbooks.
// AI Prompts: Write a C# method that reads a CSV of 12 colors (hex or R,G,B) and returns a Color[] for Aspose.Cells CustomTheme. | Generate robust error‑handling code that reports missing, extra, or malformed color lines when building a custom theme. | Provide sample code that loads a CSV palette, creates a custom theme named "BrandTheme", and saves the workbook using Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeBuilderApp
{
    // A C# utility that reads a CSV file containing exactly 12 colors (hex "#RRGGBB" or "R,G,B"), validates the entries, creates a new Workbook, applies the colors as a custom theme via Aspose.Cells, and saves the themed workbook to a specified location.
    public static class ThemeBuilder
    {
        /// <param name="csvPath">Path to the input CSV file.</param>
        /// <param name="themeName">Name of the custom theme to create.</param>
        /// <param name="outputPath">Path where the resulting workbook will be saved.</param>
        public static void BuildThemeFromCsv(string csvPath, string themeName, string outputPath)
        {
            try
            {
                // Verify CSV file exists
                if (!File.Exists(csvPath))
                    throw new FileNotFoundException($"CSV file not found: {csvPath}");

                // Read all non‑empty lines from the CSV
                string[] lines = File.ReadAllLines(csvPath);
                var colorList = new Color[12];
                int index = 0;

                foreach (string rawLine in lines)
                {
                    if (string.IsNullOrWhiteSpace(rawLine))
                        continue; // skip empty lines

                    if (index >= 12)
                        break; // we only need the first 12 colors

                    string line = rawLine.Trim();

                    // Support two formats: "R,G,B" or "#RRGGBB"
                    if (line.StartsWith("#"))
                    {
                        // Hex format
                        Color c = ColorTranslator.FromHtml(line);
                        colorList[index++] = c;
                    }
                    else
                    {
                        // CSV numeric format
                        string[] parts = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length != 3)
                            throw new FormatException($"Invalid color format at line {index + 1}: \"{line}\"");

                        int r = int.Parse(parts[0].Trim());
                        int g = int.Parse(parts[1].Trim());
                        int b = int.Parse(parts[2].Trim());

                        colorList[index++] = Color.FromArgb(r, g, b);
                    }
                }

                if (index != 12)
                    throw new InvalidOperationException($"CSV must contain exactly 12 colors. Found {index}.");

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    Directory.CreateDirectory(outDir);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Apply the custom theme
                workbook.CustomTheme(themeName, colorList);

                // Save the workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error building theme: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage
            string csvPath = "themeColors.csv";
            string themeName = "MyCsvTheme";
            string outputPath = "ThemedWorkbook.xlsx";

            try
            {
                ThemeBuilder.BuildThemeFromCsv(csvPath, themeName, outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed: {ex.Message}");
            }
        }
    }
}
