using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemePaletteReport
{
    class Program
    {
        // Helper method to print the 12 theme colors of a workbook
        static void PrintThemeColors(Workbook wb, string header)
        {
            Console.WriteLine(header);
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // ThemeColorType has values 0-12; we only need the first 12 (0-11)
                if ((int)type > 11) break;

                Color color = wb.GetThemeColor(type);
                Console.WriteLine($"  {type}: A={color.A}, R={color.R}, G={color.G}, B={color.B}");
            }
        }

        static void Main(string[] args)
        {
            // Example workbook paths – replace with actual file locations
            string[] workbookFiles = { "Workbook1.xlsx", "Workbook2.xlsx" };

            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (lifecycle rule: use constructor for loading)
                Workbook workbook = new Workbook(filePath);

                Console.WriteLine($"--- Processing: {filePath} ---");
                Console.WriteLine($"Current Theme: {workbook.Theme}");

                // Print theme colors before modification
                PrintThemeColors(workbook, "Theme colors BEFORE change:");

                // ----- Apply changes to the theme -----
                // Example: change Accent1 to Red and Accent2 to Green
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.Green);

                // Optionally, apply a full custom theme (must contain 12 colors)
                // Color[] customColors = new Color[12];
                // for (int i = 0; i < customColors.Length; i++)
                //     customColors[i] = Color.FromArgb(255, 20 * i, 30 * i, 40 * i);
                // workbook.CustomTheme("MyCustomTheme", customColors);
                // ---------------------------------------

                // Print theme colors after modification
                PrintThemeColors(workbook, "Theme colors AFTER change:");

                // Save the modified workbook (lifecycle rule: use Save method)
                string outputPath = System.IO.Path.GetFileNameWithoutExtension(filePath) + "_Modified.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Modified workbook saved as: {outputPath}");
                Console.WriteLine();
            }

            Console.WriteLine("Report generation completed.");
        }
    }
}