using System;
using System.Drawing;
using Aspose.Cells;

class ListThemeColors
{
    static void Main()
    {
        // Create a new workbook (default theme is applied)
        Workbook workbook = new Workbook();

        // Enumerate all ThemeColorType values and retrieve their current RGB colors
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // Use the Workbook.GetThemeColor method (rule-provided) to obtain the color
            Color color = workbook.GetThemeColor(type);

            // Output the theme color type and its ARGB components
            Console.WriteLine($"{type} = A:{color.A}, R:{color.R}, G:{color.G}, B:{color.B}");
        }

        // Save the workbook to satisfy the required lifecycle handling
        workbook.Save("ThemeColorsDiagnostic.xlsx");
    }
}