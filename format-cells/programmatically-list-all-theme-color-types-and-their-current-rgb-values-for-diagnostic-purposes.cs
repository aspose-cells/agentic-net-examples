// Title: List all ThemeColorType values with their ARGB values using Aspose.Cells for .NET
// Description: Creates a workbook, iterates through every ThemeColorType enum, retrieves each theme color via Workbook.GetThemeColor, prints the ARGB components, and saves the file to satisfy Aspose.Cells lifecycle rules. Ideal for diagnostics and theme verification.
// Keywords: Aspose.Cells ThemeColorType | C# get theme colors | Excel theme ARGB values | list default theme colors | Workbook.GetThemeColor example
// Common Searches: Aspose.Cells enumerate ThemeColorType | how to get RGB values of Excel theme colors .NET | list default theme colors with Aspose.Cells | retrieve ARGB of workbook theme colors | diagnostic theme colors Aspose.Cells
// Developer Intent: Obtain and display the ARGB values for every ThemeColorType in a workbook.
// Use Cases: Confirm that the default theme matches corporate color standards by printing ARGB codes. | Create a diagnostic report of all theme colors for auditing or troubleshooting. | Log theme color values when debugging unexpected color rendering in generated spreadsheets.
// AI Prompts: Generate C# code with Aspose.Cells that enumerates ThemeColorType and outputs each color's ARGB values. | Explain why saving the workbook after enumerating theme colors is required in Aspose.Cells. | Show how to format each ThemeColorType's Color as a hexadecimal string for logging purposes.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeColorDiagnostic
{
    // Creates a workbook, iterates through every ThemeColorType enum, retrieves each theme color via Workbook.GetThemeColor, prints the ARGB components, and saves the file to satisfy Aspose.Cells lifecycle rules. Ideal for diagnostics and theme verification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the default theme)
            Workbook workbook = new Workbook();

            // Loop through every ThemeColorType enum value
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the current RGB color for the theme type
                Color color = workbook.GetThemeColor(type);

                // Display the theme type and its ARGB components
                Console.WriteLine($"{type} = A:{color.A}, R:{color.R}, G:{color.G}, B:{color.B}");
            }

            // Save the workbook (no modifications are required, but saving satisfies lifecycle rules)
            workbook.Save("ThemeColorsDiagnostic.xlsx");
        }
    }
}
