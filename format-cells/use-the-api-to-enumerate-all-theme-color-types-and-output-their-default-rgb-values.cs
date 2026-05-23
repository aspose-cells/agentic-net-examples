using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (required to access theme colors)
        Workbook workbook = new Workbook();

        // Enumerate all ThemeColorType values
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // Retrieve the default theme color for the current type
            Color color = workbook.GetThemeColor(type);

            // Output the enum name and its default RGB values
            Console.WriteLine($"{type} => R:{color.R} G:{color.G} B:{color.B}");
        }

        // Save the workbook (optional, fulfills lifecycle requirement)
        workbook.Save("ThemeColors.xlsx");
    }
}