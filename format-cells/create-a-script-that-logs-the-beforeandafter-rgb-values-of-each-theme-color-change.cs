// Title: Log Before/After RGB Values of Theme Colors Using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, iterates the first 12 ThemeColorType entries, reads each original theme color, shifts its RGB components, applies the new color with SetThemeColor, prints the before and after RGB values to the console, and saves the file as ThemeColorChanges.xlsx.
// Keywords: Aspose.Cells | C# | .NET | ThemeColorType | GetThemeColor | SetThemeColor | log RGB values | theme color modification | Excel theme colors | programmatic color change | workbook save
// Common Searches: Aspose.Cells log theme color RGB values | C# get and set Excel theme colors | Iterate ThemeColorType enum Aspose.Cells | Record original and new theme colors .NET | Save workbook after changing theme colors
// Developer Intent: Capture and display the RGB values of each theme color before and after a programmatic modification in an Excel workbook.
// Use Cases: Audit branding changes by recording original and updated theme color values. | Create a design‑review report that lists before/after RGB values for all theme colors. | Debug unexpected color shifts in spreadsheets by comparing logged values.
// AI Prompts: Generate C# code that iterates ThemeColorType, logs before/after RGB values, and saves the workbook with Aspose.Cells. | Rewrite the example to write the RGB log to a CSV file instead of the console. | Explain how to revert the theme colors to their original RGB values after they have been changed.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeChangeLogger
{
    // C# example that creates a workbook, iterates the first 12 ThemeColorType entries, reads each original theme color, shifts its RGB components, applies the new color with SetThemeColor, prints the before and after RGB values to the console, and saves the file as ThemeColorChanges.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Iterate through the first 12 theme color types (Background1 to FollowedHyperlink)
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Only process the defined theme colors (indices 0‑11)
                if ((int)type > 11) continue;

                // Get the original theme color (before change)
                Color beforeColor = workbook.GetThemeColor(type);

                // Define a new color – for demonstration we shift each RGB component by +50 (wrap around 255)
                Color afterColor = Color.FromArgb(
                    (beforeColor.R + 50) % 256,
                    (beforeColor.G + 50) % 256,
                    (beforeColor.B + 50) % 256);

                // Apply the new theme color (lifecycle rule: modify)
                workbook.SetThemeColor(type, afterColor);

                // Log the before‑and‑after RGB values
                Console.WriteLine($"{type}: Before = {beforeColor.R},{beforeColor.G},{beforeColor.B} | After = {afterColor.R},{afterColor.G},{afterColor.B}");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ThemeColorChanges.xlsx");
        }
    }
}
