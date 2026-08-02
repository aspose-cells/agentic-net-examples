// Title: List all Aspose.Cells ThemeColorType values with their RGB (ARGB) values in C#
// Description: Creates a new Workbook, iterates through every ThemeColorType enum member, uses Workbook.GetThemeColor to fetch the current theme color, prints each type with its ARGB components, and saves the file for lifecycle compliance.
// Keywords: Aspose.Cells | ThemeColorType | GetThemeColor | RGB values | ARGB values | C# | Excel theme colors | list theme colors | diagnostic report | color enumeration
// Common Searches: Aspose.Cells get theme color RGB | C# list ThemeColorType enum values | How to read Excel theme colors with Aspose | Retrieve ARGB values of workbook theme colors
// Developer Intent: Fetch and show the RGB/ARGB values for every ThemeColorType defined in a workbook.
// Use Cases: Verify that a workbook's theme matches corporate branding standards. | Generate a diagnostic sheet of all theme colors for debugging custom themes. | Compare theme palettes across multiple workbooks to ensure visual consistency.
// AI Prompts: Write a C# method that returns a Dictionary<ThemeColorType, Color> containing all theme colors from an Aspose.Cells Workbook. | Provide code to export ThemeColorType names and their ARGB values to a CSV file using Aspose.Cells. | Create a PowerShell script that calls a compiled C# program to list Excel theme colors and logs the output to a text file.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeColorDiagnostic
{
    // Creates a new Workbook, iterates through every ThemeColorType enum member, uses Workbook.GetThemeColor to fetch the current theme color, prints each type with its ARGB components, and saves the file for lifecycle compliance.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default theme is applied)
            Workbook workbook = new Workbook();

            // Iterate through all ThemeColorType enum values
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the current RGB color for the theme type
                Color color = workbook.GetThemeColor(type);

                // Output the theme type name and its ARGB components
                Console.WriteLine($"{type} => A={color.A}, R={color.R}, G={color.G}, B={color.B}");
            }

            // Save the workbook (required by lifecycle rules)
            workbook.Save("ThemeColorDiagnostic.xlsx");
        }
    }
}
