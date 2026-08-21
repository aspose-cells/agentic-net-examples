// Title: List All ThemeColorType Enum Values and Their Default RGB Colors with Aspose.Cells for .NET
// Description: Shows how to create a Workbook, loop through the ThemeColorType enumeration, retrieve each default theme color using Workbook.GetThemeColor, and print the enum name with its R, G, B components. The sample also demonstrates optional workbook saving.
// Keywords: Aspose.Cells | ThemeColorType | GetThemeColor | default theme colors | RGB values | C# | .NET | Excel theme palette | enumerate theme colors | workbook theme colors | color enumeration
// Common Searches: Aspose.Cells list ThemeColorType values | Get default RGB for Excel theme colors C# | Workbook.GetThemeColor example | enumerate Excel theme colors with Aspose | ThemeColorType enum RGB Aspose.Cells
// Developer Intent: Retrieve and display the built‑in RGB values for every ThemeColorType in an Aspose.Cells workbook.
// Use Cases: Create a reference table of Excel theme colors for UI design or documentation. | Validate that custom theme colors match the default palette in a workbook. | Generate a color legend for reports by iterating over ThemeColorType and outputting RGB values.
// AI Prompts: Write a method that returns a Dictionary<ThemeColorType, Color> containing the default theme colors using Aspose.Cells. | Modify the example to export ThemeColorType names and their RGB values to a CSV file. | Explain how Workbook.GetThemeColor determines the color and which theme is applied when a new workbook is created.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeColorDemo
{
    // Shows how to create a Workbook, loop through the ThemeColorType enumeration, retrieve each default theme color using Workbook.GetThemeColor, and print the enum name with its R, G, B components. The sample also demonstrates optional workbook saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default theme is applied)
            Workbook workbook = new Workbook();

            // Iterate through all ThemeColorType enum values
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Get the default theme color for the current type
                Color color = workbook.GetThemeColor(type);

                // Output the enum name and its RGB components
                Console.WriteLine($"{type}: R={color.R}, G={color.G}, B={color.B}");
            }

            // Optionally save the workbook if you want to inspect the theme in Excel
            workbook.Save("ThemeColorsDemo.xlsx");
        }
    }
}
