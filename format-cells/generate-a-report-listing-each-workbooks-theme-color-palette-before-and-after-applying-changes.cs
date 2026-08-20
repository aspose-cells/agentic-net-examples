// Title: Aspose.Cells C# console app to list and compare Excel workbook theme colors before and after changes
// Description: A C# console program that loads one or more Excel files, reads the 12 theme colors (excluding StyleColor), prints the original palette, modifies Accent1 and Accent2, prints the updated palette, and saves the workbook with a new name. Ideal for auditing and documenting theme‑color updates.
// Keywords: Aspose.Cells | C# | .NET | Excel theme colors | GetThemeColor | SetThemeColor | theme palette report | before after comparison | console application | workbook audit
// Common Searches: how to retrieve Excel theme colors using Aspose.Cells C# | Aspose.Cells change theme accent colors programmatically | list workbook theme palette before and after modification | C# code to compare Excel theme colors
// Developer Intent: Generate a console‑based report that shows each workbook’s theme color palette, applies specific color changes, and displays the before/after values.
// Use Cases: Validate that corporate branding colors are applied across multiple workbooks. | Document theme‑color changes for compliance or version‑control purposes. | Automate bulk updates of Excel theme accents while preserving a change log.
// AI Prompts: Create code to export the before‑and‑after theme color data to a CSV file instead of console output. | Add functionality to revert all theme colors to their original values after processing each workbook. | Write a method that accepts a dictionary of ThemeColorType‑Color pairs and applies them in a single loop.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeReport
{
    // A C# console program that loads one or more Excel files, reads the 12 theme colors (excluding StyleColor), prints the original palette, modifies Accent1 and Accent2, prints the updated palette, and saves the workbook with a new name. Ideal for auditing and documenting theme‑color updates.
    class Program
    {
        // Retrieves the 12 theme colors of a workbook (excluding StyleColor)
        static Dictionary<ThemeColorType, Color> GetThemeColors(Workbook wb)
        {
            var dict = new Dictionary<ThemeColorType, Color>();
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                if (type == ThemeColorType.StyleColor) continue;
                dict[type] = wb.GetThemeColor(type);
            }
            return dict;
        }

        // Prints theme colors to the console
        static void PrintThemeColors(string header, Dictionary<ThemeColorType, Color> colors)
        {
            Console.WriteLine(header);
            foreach (var kvp in colors)
            {
                Color c = kvp.Value;
                Console.WriteLine($"{kvp.Key}: A={c.A}, R={c.R}, G={c.G}, B={c.B}");
            }
            Console.WriteLine();
        }

        static void Main(string[] args)
        {
            string[] workbookFiles = new string[]
            {
                "Input1.xlsx",
                "Input2.xlsx"
                // Add more file names as needed
            };

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}. Skipping.");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);
                    Console.WriteLine($"Processing workbook: {filePath}");
                    Console.WriteLine($"Current theme name: {workbook.Theme}");
                    Console.WriteLine();

                    // Theme colors before changes
                    var beforeColors = GetThemeColors(workbook);
                    PrintThemeColors("Theme colors BEFORE changes:", beforeColors);

                    // ----- Apply changes -----
                    workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);
                    workbook.SetThemeColor(ThemeColorType.Accent2, Color.Green);

                    // Capture theme colors after changes
                    var afterColors = GetThemeColors(workbook);
                    PrintThemeColors("Theme colors AFTER changes:", afterColors);

                    // Save the modified workbook
                    string outputPath = Path.GetFileNameWithoutExtension(filePath) + "_Modified.xlsx";
                    workbook.Save(outputPath);
                    Console.WriteLine($"Modified workbook saved as: {outputPath}");
                    Console.WriteLine(new string('-', 50));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
