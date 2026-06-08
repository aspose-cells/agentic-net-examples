using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeColorChangeLogger
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define new colors for each theme type (sample colors)
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
                Color.FromArgb(255, 150, 255), // Accent5
                Color.FromArgb(150, 255, 255), // Accent6
                Color.FromArgb(100, 100, 255), // Hyperlink
                Color.FromArgb(255, 100, 100)  // FollowedHyperlink
            };

            // Iterate over each ThemeColorType (0-11)
            ThemeColorType[] themeTypes = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));
            for (int i = 0; i < themeTypes.Length; i++)
            {
                ThemeColorType type = themeTypes[i];
                // Skip StyleColor (value 12) which is not a real theme slot
                if (type == ThemeColorType.StyleColor) continue;

                // Get the original color
                Color before = workbook.GetThemeColor(type);

                // Apply the new color (ensure we have a corresponding entry)
                Color afterColor = newColors[(int)type];
                workbook.SetThemeColor(type, afterColor);

                // Get the updated color
                Color after = workbook.GetThemeColor(type);

                // Log before and after RGB values
                Console.WriteLine($"{type}: Before = {before.ToArgb()}, After = {after.ToArgb()}");
            }

            // Save the workbook (theme changes are persisted)
            workbook.Save("ThemeColorChangeLog.xlsx");
        }
    }
}