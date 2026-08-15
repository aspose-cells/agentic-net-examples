// Title: Benchmark applying a custom 12‑color theme to a large Aspose.Cells workbook (C#)
// Description: Creates a 10,000‑row by 10‑column worksheet, defines a full 12‑color theme, measures the elapsed time with Stopwatch while calling workbook.CustomTheme, outputs the duration in milliseconds, and saves the file. Demonstrates end‑to‑end performance testing for theme updates in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | custom theme performance | theme color update benchmark | workbook.CustomTheme timing | large workbook Excel | stopwatch performance test | 12 theme colors | Excel theme update speed | performance testing Aspose.Cells
// Common Searches: how long does workbook.CustomTheme take on a 10k row Excel file | Aspose.Cells theme update performance C# | measure time to apply custom theme in Aspose.Cells | benchmark theme color changes in large spreadsheet | speed of updating all theme colors with Aspose.Cells
// Developer Intent: The developer wants to measure and benchmark the execution time of applying a full custom theme (12 colors) to a large workbook using Aspose.Cells.
// Use Cases: Validate that theme changes meet performance SLAs for real‑time report generation | Compare the impact of different theme palettes on processing time in batch spreadsheet workflows | Identify bottlenecks before optimizing large‑scale Excel export services
// AI Prompts: Generate C# code that logs the time for each individual theme color update with Aspose.Cells. | Suggest optimization techniques to reduce workbook.CustomTheme latency for workbooks with millions of cells. | Create an xUnit test that asserts the custom theme update completes within a configurable time threshold.

using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Cells;

namespace ThemeColorUpdateTiming
{
    // Creates a 10,000‑row by 10‑column worksheet, defines a full 12‑color theme, measures the elapsed time with Stopwatch while calling workbook.CustomTheme, outputs the duration in milliseconds, and saves the file. Demonstrates end‑to‑end performance testing for theme updates in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate a large worksheet with dummy data (e.g., 10,000 rows x 10 columns)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Prepare new theme colors (12 colors required)
            Color[] newThemeColors = new Color[12];
            newThemeColors[0] = Color.FromArgb(255, 255, 200, 200); // Background1
            newThemeColors[1] = Color.FromArgb(255, 200, 255, 200); // Text1
            newThemeColors[2] = Color.FromArgb(255, 200, 200, 255); // Background2
            newThemeColors[3] = Color.FromArgb(255, 255, 255, 200); // Text2
            newThemeColors[4] = Color.FromArgb(255, 255, 150, 150); // Accent1
            newThemeColors[5] = Color.FromArgb(255, 150, 255, 150); // Accent2
            newThemeColors[6] = Color.FromArgb(255, 150, 150, 255); // Accent3
            newThemeColors[7] = Color.FromArgb(255, 255, 255, 150); // Accent4
            newThemeColors[8] = Color.FromArgb(255, 255, 150, 255); // Accent5
            newThemeColors[9] = Color.FromArgb(255, 150, 255, 255); // Accent6
            newThemeColors[10] = Color.FromArgb(255, 0, 0, 255);    // Hyperlink
            newThemeColors[11] = Color.FromArgb(255, 128, 0, 128); // Followed Hyperlink

            // Measure the time taken to apply the custom theme
            Stopwatch sw = Stopwatch.StartNew();

            // Apply the custom theme (updates all 12 theme colors at once)
            workbook.CustomTheme("PerformanceTestTheme", newThemeColors);

            sw.Stop();

            Console.WriteLine($"Time taken to update all theme colors: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ThemeUpdateTiming.xlsx");
        }
    }
}
