using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default theme is applied)
        Workbook workbook = new Workbook();

        // Enumerate all ThemeColorType values
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // Retrieve the default color for the current theme type
            Color color = workbook.GetThemeColor(type);

            // Output the theme type and its RGB components
            Console.WriteLine($"{type} => R:{color.R} G:{color.G} B:{color.B}");
        }

        // Save the workbook (optional, just to follow lifecycle rules)
        workbook.Save("ThemeColors.xlsx");
    }
}