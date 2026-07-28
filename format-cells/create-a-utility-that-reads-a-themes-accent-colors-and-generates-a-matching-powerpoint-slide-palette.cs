// Title: C# – Extract Excel Theme Accent Colors with Aspose.Cells for PowerPoint Palette
// Description: Loads an Excel workbook using Aspose.Cells, reads the six theme accent colors (Accent1‑Accent6) via GetThemeColor, displays RGBA values, and optionally saves them to a CSV file that can be imported into PowerPoint to create a matching slide palette.
// Keywords: Aspose.Cells GetThemeColor C# | read Excel theme accent colors | export theme colors to CSV | Excel to PowerPoint color palette | C# extract workbook theme colors
// Common Searches: how to get Excel theme accent colors with Aspose.Cells | export Excel theme colors to CSV in C# | use Excel theme colors for PowerPoint palette | Aspose.Cells GetThemeColor example | C# read workbook theme colors
// Developer Intent: Read an Excel workbook’s theme accent colors and export them for reuse in a PowerPoint slide palette.
// Use Cases: Synchronize branding by extracting the six accent colors from an Excel theme and applying them to PowerPoint slides. | Generate a CSV file of RGBA values for designers to import into presentation templates. | Quickly verify theme colors during development by printing RGBA components to the console.
// AI Prompts: Write C# code that uses Aspose.Cells to retrieve all six theme accent colors from an Excel file and returns them as a List<System.Drawing.Color>. | Show how to modify the example to output the accent colors in HEX format and save them to a JSON file for PowerPoint integration. | Explain how to handle workbooks with custom themes or missing accent definitions when calling GetThemeColor.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeAccentReader
{
    // Loads an Excel workbook using Aspose.Cells, reads the six theme accent colors (Accent1‑Accent6) via GetThemeColor, displays RGBA values, and optionally saves them to a CSV file that can be imported into PowerPoint to create a matching slide palette.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ThemeAccentReader <excel-file-path>");
                return;
            }

            string excelPath = args[0];

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(excelPath);

            // Retrieve accent colors from the workbook theme
            Color[] accentColors = new Color[6];
            accentColors[0] = workbook.GetThemeColor(ThemeColorType.Accent1);
            accentColors[1] = workbook.GetThemeColor(ThemeColorType.Accent2);
            accentColors[2] = workbook.GetThemeColor(ThemeColorType.Accent3);
            accentColors[3] = workbook.GetThemeColor(ThemeColorType.Accent4);
            accentColors[4] = workbook.GetThemeColor(ThemeColorType.Accent5);
            accentColors[5] = workbook.GetThemeColor(ThemeColorType.Accent6);

            // Output the accent colors – these can be used to build a matching PowerPoint slide palette
            Console.WriteLine("Accent colors extracted from theme:");
            for (int i = 0; i < accentColors.Length; i++)
            {
                Color c = accentColors[i];
                Console.WriteLine($"Accent{i + 1}: R={c.R}, G={c.G}, B={c.B}, A={c.A}");
            }

            // OPTIONAL: Save the colors to a simple CSV file for external use
            string csvPath = System.IO.Path.ChangeExtension(excelPath, ".csv");
            using (var writer = new System.IO.StreamWriter(csvPath))
            {
                writer.WriteLine("Accent,Red,Green,Blue,Alpha");
                for (int i = 0; i < accentColors.Length; i++)
                {
                    Color c = accentColors[i];
                    writer.WriteLine($"Accent{i + 1},{c.R},{c.G},{c.B},{c.A}");
                }
            }

            Console.WriteLine($"Accent colors have been written to: {csvPath}");
        }
    }
}
