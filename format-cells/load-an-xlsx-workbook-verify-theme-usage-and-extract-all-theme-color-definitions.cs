using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeExtraction
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook has a theme assigned
            string themeName = workbook.Theme;
            Console.WriteLine($"Workbook Theme: {(string.IsNullOrEmpty(themeName) ? "None" : themeName)}");

            // Extract all theme color definitions
            Console.WriteLine("Theme Colors:");
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Get the color for the current theme type (rule: GetThemeColor)
                Color color = workbook.GetThemeColor(type);
                Console.WriteLine($"{type}: A={color.A}, R={color.R}, G={color.G}, B={color.B}");
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("OutputWorkbook.xlsx"); // lifecycle rule: save
        }
    }
}