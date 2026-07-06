using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Cells;

namespace ThemeColorUpdateTiming
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate a large worksheet to simulate a heavy workbook
            Worksheet sheet = workbook.Worksheets[0];
            const int rows = 10000;
            const int cols = 10;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r}C{c}");
                }
            }

            // Prepare random colors for the 12 theme slots
            Random rnd = new Random();
            Color[] newThemeColors = new Color[12];
            for (int i = 0; i < newThemeColors.Length; i++)
            {
                newThemeColors[i] = Color.FromArgb(255, rnd.Next(256), rnd.Next(256), rnd.Next(256));
            }

            // Measure the time taken to update all theme colors
            Stopwatch sw = Stopwatch.StartNew();

            // Update each theme color using SetThemeColor
            ThemeColorType[] themeTypes = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));
            for (int i = 0; i < themeTypes.Length && i < newThemeColors.Length; i++)
            {
                workbook.SetThemeColor(themeTypes[i], newThemeColors[i]);
            }

            sw.Stop();
            Console.WriteLine($"Time taken to update all theme colors: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (lifecycle rule)
            workbook.Save("LargeWorkbook_WithUpdatedTheme.xlsx");
        }
    }
}