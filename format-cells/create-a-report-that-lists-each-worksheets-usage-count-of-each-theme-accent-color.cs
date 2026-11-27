// Title: Generate a per‑worksheet Excel theme accent color usage report with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# count how many cells use each theme accent color in a workbook | Create a report of theme accent color occurrences per worksheet using Aspose.Cells | C# code to list usage of Accent1‑Accent6 colors across Excel sheets with Aspose | How to iterate cells and tally theme accent colors in Aspose.Cells | Generate an Excel file that summarizes theme accent color counts per sheet in .NET
// Tags: Aspose.Cells GetThemeColor palette extraction | C# count theme accent colors in Excel | worksheet level theme color statistics | auto‑generate accent count workbook | scan cell style colors with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, retrieves the six theme accent colors via GetThemeColor, scans every cell in each worksheet’s used range, compares the cell’s font, foreground, and background colors to the accent palette, tallies the occurrences for each accent, and writes the worksheet name together with the six accent counts into a new workbook named "ThemeAccentReport.xlsx".
class ThemeAccentUsageReport
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "ThemeAccentReport.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Theme Accent Usage";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            for (int i = 0; i < 6; i++)
            {
                reportSheet.Cells[0, i + 1].PutValue($"Accent{i + 1} Count");
            }

            int reportRow = 1;

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Retrieve theme accent colors using GetThemeColor
                Color[] accentColors = new Color[6];
                try
                {
                    accentColors[0] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent1);
                    accentColors[1] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent2);
                    accentColors[2] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent3);
                    accentColors[3] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent4);
                    accentColors[4] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent5);
                    accentColors[5] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent6);
                }
                catch
                {
                    // If theme colors cannot be retrieved, leave them empty
                }

                int[] accentCounts = new int[6];

                // Determine the used range to limit iteration
                Aspose.Cells.Range usedRange = ws.Cells.MaxDisplayRange;
                int firstRow = usedRange.FirstRow;
                int firstColumn = usedRange.FirstColumn;
                int totalRows = usedRange.RowCount;
                int totalColumns = usedRange.ColumnCount;

                // Iterate through each cell in the used range
                for (int r = firstRow; r < firstRow + totalRows; r++)
                {
                    for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                    {
                        Cell cell = ws.Cells[r, c];
                        if (cell == null) continue;

                        Style style = cell.GetStyle();

                        // Check Font color
                        IncrementAccentCount(style.Font.Color, accentColors, accentCounts);

                        // Check Foreground (pattern) color
                        IncrementAccentCount(style.ForegroundColor, accentColors, accentCounts);

                        // Check Background color
                        IncrementAccentCount(style.BackgroundColor, accentColors, accentCounts);
                    }
                }

                // Write results to the report sheet
                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                for (int i = 0; i < 6; i++)
                {
                    reportSheet.Cells[reportRow, i + 1].PutValue(accentCounts[i]);
                }
                reportRow++;
            }

            // Auto-fit columns for better readability
            reportSheet.AutoFitColumns();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the report workbook
            reportWorkbook.Save(outputPath);
            Console.WriteLine($"Report saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to increment the appropriate accent counter if the color matches an accent
    private static void IncrementAccentCount(Color color, Color[] accentColors, int[] accentCounts)
    {
        // Ignore automatic or empty colors
        if (color.IsEmpty) return;

        for (int i = 0; i < accentColors.Length; i++)
        {
            if (!accentColors[i].IsEmpty && color.ToArgb() == accentColors[i].ToArgb())
            {
                accentCounts[i]++;
                break;
            }
        }
    }
}
