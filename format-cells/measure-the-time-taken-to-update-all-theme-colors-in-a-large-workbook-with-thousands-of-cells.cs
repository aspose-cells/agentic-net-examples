using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Cells;

class ThemeUpdatePerformance
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate a large number of cells to simulate a big workbook
        int rows = 5000;
        int cols = 20;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                sheet.Cells[r, c].PutValue($"R{r}C{c}");
            }
        }

        // Prepare new theme colors (example values)
        Color[] newColors = new Color[12];
        newColors[0] = Color.FromArgb(255, 200, 200, 200); // Background1
        newColors[1] = Color.FromArgb(255, 0, 0, 0);       // Text1
        newColors[2] = Color.FromArgb(255, 230, 230, 230); // Background2
        newColors[3] = Color.FromArgb(255, 80, 80, 80);    // Text2
        newColors[4] = Color.FromArgb(255, 255, 0, 0);     // Accent1
        newColors[5] = Color.FromArgb(255, 0, 255, 0);     // Accent2
        newColors[6] = Color.FromArgb(255, 0, 0, 255);     // Accent3
        newColors[7] = Color.FromArgb(255, 255, 255, 0);   // Accent4
        newColors[8] = Color.FromArgb(255, 255, 0, 255);   // Accent5
        newColors[9] = Color.FromArgb(255, 0, 255, 255);   // Accent6
        newColors[10] = Color.FromArgb(255, 0, 0, 255);    // Hyperlink
        newColors[11] = Color.FromArgb(255, 128, 0, 128);  // FollowedHyperlink

        // Measure time to update all theme colors
        Stopwatch sw = Stopwatch.StartNew();

        // Update each theme color using SetThemeColor
        ThemeColorType[] types = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));
        for (int i = 0; i < types.Length && i < newColors.Length; i++)
        {
            workbook.SetThemeColor(types[i], newColors[i]);
        }

        sw.Stop();

        Console.WriteLine($"Time to update all theme colors: {sw.ElapsedMilliseconds} ms");

        // Save the workbook (optional, to ensure changes are persisted)
        workbook.Save("LargeWorkbook_ThemeUpdated.xlsx");
    }
}