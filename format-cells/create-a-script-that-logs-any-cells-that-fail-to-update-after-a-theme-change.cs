// Title: C# Aspose.Cells example: Log cells that don't update after a theme change
// Description: A complete C# script that creates a workbook, applies a custom theme, scans every used cell, and logs those whose Font.ThemeColor is missing or not an accent color. The console output shows the cell address and reason, and the workbook is saved for further review.
// Keywords: Aspose.Cells C# theme change | detect cells without ThemeColor | log unupdated cells Aspose | custom theme verification | font ThemeColor detection | Excel theme audit C# | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells log cells after theme change | C# find cells not using theme colors | detect hard‑coded font colors in Excel with Aspose | audit workbook after applying custom theme | list cells unchanged by theme in Aspose.Cells
// Developer Intent: Find and record cells whose formatting does not reflect a newly applied custom theme.
// Use Cases: Validate brand‑compliant styling before publishing a report. | Generate a warning list for hard‑coded colors after a theme migration. | Audit large workbooks to pinpoint cells needing manual style updates.
// AI Prompts: Create a method that returns cell addresses where Font.ThemeColor is null or not an accent after applying a custom theme in Aspose.Cells. | Show C# code to export the unupdated cell list to a CSV file instead of the console. | Explain how to extend the detection to include background fill ThemeColor checks in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeChangeLogger
{
    // Custom class to hold information about cells that did not update after a theme change
    // A complete C# script that creates a workbook, applies a custom theme, scans every used cell, and logs those whose Font.ThemeColor is missing or not an accent color. The console output shows the cell address and reason, and the workbook is saved for further review.
    public class UnupdatedCellInfo
    {
        public string CellName { get; set; }
        public string Reason { get; set; }
    }

    public class ThemeChangeLogger
    {
        public static void Run()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with values
            sheet.Cells["A1"].PutValue("Default Theme");
            sheet.Cells["B2"].PutValue("Uses Accent1");
            sheet.Cells["C3"].PutValue("Uses Accent2");
            sheet.Cells["D4"].PutValue("No Theme Color");

            // Apply theme colors to some cells using ThemeColor
            // Cell B2 -> Accent1
            Style accent1Style = workbook.CreateStyle();
            accent1Style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            sheet.Cells["B2"].SetStyle(accent1Style);

            // Cell C3 -> Accent2
            Style accent2Style = workbook.CreateStyle();
            accent2Style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);
            sheet.Cells["C3"].SetStyle(accent2Style);

            // Cell D4 uses a direct color (no theme)
            Style directStyle = workbook.CreateStyle();
            directStyle.Font.Color = Color.Black;
            sheet.Cells["D4"].SetStyle(directStyle);

            // -------------------------------------------------
            // 2. Define a custom theme (12 colors as required)
            // -------------------------------------------------
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(50, 50, 50),    // Text2
                Color.FromArgb(255, 0, 0),     // Accent1 (Red)
                Color.FromArgb(0, 255, 0),     // Accent2 (Green)
                Color.FromArgb(0, 0, 255),     // Accent3 (Blue)
                Color.FromArgb(255, 255, 0),   // Accent4 (Yellow)
                Color.FromArgb(255, 0, 255),   // Accent5 (Magenta)
                Color.FromArgb(0, 255, 255),   // Accent6 (Cyan)
                Color.FromArgb(0, 0, 128),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };

            // -------------------------------------------------
            // 3. Apply the custom theme
            // -------------------------------------------------
            workbook.CustomTheme("MyCustomTheme", customColors);

            // -------------------------------------------------
            // 4. After theme change, detect cells that did NOT
            //    use a theme color (they will not reflect the new theme)
            // -------------------------------------------------
            List<UnupdatedCellInfo> unupdatedCells = new List<UnupdatedCellInfo>();

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Retrieve the style of the current cell
                Style cellStyle = cell.GetStyle();

                // If the Font.ThemeColor is null, the cell uses a direct color
                // and therefore will not be affected by the theme change.
                if (cellStyle.Font.ThemeColor == null)
                {
                    unupdatedCells.Add(new UnupdatedCellInfo
                    {
                        CellName = cell.Name,
                        Reason = "Font does not use ThemeColor"
                    });
                }
                else
                {
                    // Optional: verify that the ThemeColor type matches one of the
                    // custom theme's accent colors (Accent1‑Accent6). If it does not,
                    // it may also be considered as not updated.
                    ThemeColorType type = cellStyle.Font.ThemeColor.ColorType;
                    if (type != ThemeColorType.Accent1 &&
                        type != ThemeColorType.Accent2 &&
                        type != ThemeColorType.Accent3 &&
                        type != ThemeColorType.Accent4 &&
                        type != ThemeColorType.Accent5 &&
                        type != ThemeColorType.Accent6)
                    {
                        unupdatedCells.Add(new UnupdatedCellInfo
                        {
                            CellName = cell.Name,
                            Reason = $"Font uses ThemeColor type {type}, which is not an accent color"
                        });
                    }
                }
            }

            // -------------------------------------------------
            // 5. Log the results
            // -------------------------------------------------
            Console.WriteLine("Cells that failed to update after the theme change:");
            if (unupdatedCells.Count == 0)
            {
                Console.WriteLine("  None – all cells use theme colors.");
            }
            else
            {
                foreach (var info in unupdatedCells)
                {
                    Console.WriteLine($"  {info.CellName}: {info.Reason}");
                }
            }

            // -------------------------------------------------
            // 6. Save the workbook (lifecycle rule)
            // -------------------------------------------------
            workbook.Save("ThemeChangeLogDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main(string[] args)
        {
            ThemeChangeLogger.Run();
        }
    }
}
