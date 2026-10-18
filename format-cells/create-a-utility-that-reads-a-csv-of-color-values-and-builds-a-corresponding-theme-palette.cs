// Title: Create a C# Aspose.Cells program that reads a CSV of hex color codes and generates an Excel workbook with a visual palette
// Common Searches: how to import hex color values from a CSV file and set cell background colors using Aspose.Cells in C# | C# example for building an Excel palette worksheet from a list of HTML color codes with Aspose.Cells | reading a CSV of color codes and applying them to cells in an XLSX workbook using Aspose.Cells | convert CSV hex colors to System.Drawing.Color and generate a colored Excel file with Aspose.Cells | Aspose.Cells C# tutorial for creating a custom theme palette from external color data
// Tags: read CSV hex color values C# | set cell background Aspose.Cells | create palette worksheet Aspose.Cells | HTML hex to System.Drawing.Color conversion | save Excel workbook as XLSX Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace ThemePaletteBuilder
{
    // The utility reads a CSV file containing HTML hex color strings, converts each entry to a System.Drawing.Color, creates a new Workbook with Aspose.Cells, writes the hex values into column A of a worksheet, applies the corresponding background color to each cell, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input CSV path and output Excel file path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: ThemePaletteBuilder <colors.csv> <output.xlsx>");
                return;
            }

            string csvPath = args[0];
            string outputPath = args[1];

            // Verify CSV file exists
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
                return;
            }

            try
            {
                // Read all lines from the CSV file. Each line should contain a color value (e.g., #FF5733)
                string[] lines = File.ReadAllLines(csvPath);
                Color[] colors = new Color[lines.Length];

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();

                    // Skip empty lines
                    if (string.IsNullOrEmpty(line))
                    {
                        colors[i] = Color.Empty;
                        continue;
                    }

                    try
                    {
                        // Convert HTML hex color to System.Drawing.Color
                        colors[i] = ColorTranslator.FromHtml(line);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Invalid color value at line {i + 1}: '{line}'. Error: {ex.Message}");
                        return;
                    }
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet to display the colors (since direct theme manipulation is unavailable in this version)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Palette";

                // Populate the worksheet with color samples
                for (int i = 0; i < colors.Length; i++)
                {
                    // Write the hex value (or empty) in column A
                    Cell cell = sheet.Cells[i, 0];
                    cell.PutValue(string.IsNullOrEmpty(lines[i]) ? string.Empty : lines[i]);

                    // Apply background color if valid
                    if (colors[i] != Color.Empty)
                    {
                        Style style = cell.GetStyle();
                        style.ForegroundColor = colors[i];
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Palette workbook created and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
