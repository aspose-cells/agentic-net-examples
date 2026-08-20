// Title: C# Example: Log Before‑and‑After RGB Values for Workbook Theme Colors with Aspose.Cells
// Description: This Aspose.Cells for .NET sample creates a workbook, defines a 12‑color palette, iterates the ThemeColorType enum (excluding StyleColor), reads each original theme color with GetThemeColor, applies a new color via SetThemeColor, writes the before‑and‑after RGB values to the console, and saves the file as ThemeColorChangesLog.xlsx.
// Keywords: Aspose.Cells C# theme color | GetThemeColor | SetThemeColor | ThemeColorType enumeration | log RGB values Aspose.Cells | Excel theme color change | C# workbook theme audit | Aspose.Cells example .NET
// Common Searches: how to log theme color changes Aspose.Cells | retrieve original theme colors before modification C# | iterate ThemeColorType enum Aspose.Cells | save workbook after changing theme colors .NET | C# console output RGB values Excel theme
// Developer Intent: Record the original and updated RGB components for every workbook theme color when applying a new palette using Aspose.Cells.
// Use Cases: Create an audit trail of theme‑color modifications for compliance or QA. | Debug visual differences in generated Excel reports by comparing color values. | Maintain version‑controlled logs of theme updates across automated report pipelines.
// AI Prompts: Generate C# code that logs before and after RGB values for each ThemeColorType when changing a workbook's theme with Aspose.Cells. | Explain how to exclude the StyleColor enum value while looping through ThemeColorType. | Show how to write the theme‑color change log to a CSV file instead of the console.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeColorChangeLogger
{
    // This Aspose.Cells for .NET sample creates a workbook, defines a 12‑color palette, iterates the ThemeColorType enum (excluding StyleColor), reads each original theme color with GetThemeColor, applies a new color via SetThemeColor, writes the before‑and‑after RGB values to the console, and saves the file as ThemeColorChangesLog.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Define new colors for each theme type (must be 12 entries)
            Color[] newColors = new Color[]
            {
                Color.FromArgb(255, 255, 200), // Background1
                Color.FromArgb(200, 255, 255), // Text1
                Color.FromArgb(255, 200, 255), // Background2
                Color.FromArgb(200, 200, 255), // Text2
                Color.FromArgb(255, 150, 150), // Accent1
                Color.FromArgb(150, 255, 150), // Accent2
                Color.FromArgb(150, 150, 255), // Accent3
                Color.FromArgb(255, 255, 150), // Accent4
                Color.FromArgb(150, 255, 255), // Accent5
                Color.FromArgb(255, 150, 255), // Accent6
                Color.FromArgb(100, 100, 255), // Hyperlink
                Color.FromArgb(255, 100, 100)  // FollowedHyperlink
            };

            // Iterate over each ThemeColorType (0 to 11)
            ThemeColorType[] types = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));
            for (int i = 0; i < types.Length; i++)
            {
                ThemeColorType type = types[i];
                // Skip StyleColor (value 12) which is not a theme color
                if (type == ThemeColorType.StyleColor) continue;

                // Get the original color
                Color before = workbook.GetThemeColor(type);

                // Apply the new color
                workbook.SetThemeColor(type, newColors[i]);

                // Get the updated color
                Color after = workbook.GetThemeColor(type);

                // Log before and after RGB values
                Console.WriteLine($"{type}: Before = ({before.R}, {before.G}, {before.B}) -> After = ({after.R}, {after.G}, {after.B})");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("ThemeColorChangesLog.xlsx");
        }
    }
}
