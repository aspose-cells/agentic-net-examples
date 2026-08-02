using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeExtractor
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Verify that a theme is applied and display its name
            Console.WriteLine("Workbook Theme: " + workbook.Theme);

            // Extract and display all theme color definitions
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                Color themeColor = workbook.GetThemeColor(type);
                Console.WriteLine($"{type}: A={themeColor.A}, R={themeColor.R}, G={themeColor.G}, B={themeColor.B}");
            }

            // Optionally, save the workbook (demonstrates lifecycle compliance)
            string outputPath = "output_with_verified_theme.xlsx";
            workbook.Save(outputPath);
        }
    }
}