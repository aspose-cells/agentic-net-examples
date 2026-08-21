// Title: C# – Log Cells That Fail to Update After a Theme Change with Aspose.Cells
// Description: Demonstrates how to create a workbook, apply a ThemeColor‑based style to selected cells, switch the workbook to a custom theme, and then scan the recorded cells to detect which ones lost their ThemeColor reference, logging success or failure before optionally saving the file.
// Keywords: Aspose.Cells | C# | .NET | Excel theme change | ThemeColor detection | cell style audit | logging failed updates | custom workbook theme | automated Excel validation
// Common Searches: Aspose.Cells log cells after theme change C# | detect cells that lose ThemeColor in Excel .NET | verify theme consistency after applying custom theme | C# code to track cells not updating with new Excel theme
// Developer Intent: Find and record worksheet cells whose ThemeColor reference is lost when a new theme is applied.
// Use Cases: Quality‑check financial or marketing reports to ensure branding colors persist after a corporate theme update. | Batch‑process workbooks to flag cells that need manual style correction after migrating to a new color scheme. | Integrate into CI pipelines to automatically verify that theme‑dependent formatting remains intact across builds.
// AI Prompts: Generate a reusable method that returns a list of cell addresses with a null Font.ThemeColor after a theme switch using Aspose.Cells. | Create code to export the failure log (address and status) to CSV for downstream analysis. | Write a unit test that confirms ThemeChangeLogger correctly identifies cells that lose ThemeColor when a custom theme is applied.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeChangeLogger
{
    // This class demonstrates how to log cells that do not reflect a theme change.
    // Demonstrates how to create a workbook, apply a ThemeColor‑based style to selected cells, switch the workbook to a custom theme, and then scan the recorded cells to detect which ones lost their ThemeColor reference, logging success or failure before optionally saving the file.
    public class ThemeChangeLogger
    {
        public static void Run()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample data.
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style that uses a theme color (Accent1) and apply it to several cells.
            Style themeStyle = workbook.CreateStyle();
            themeStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            themeStyle.Font.Size = 12;
            themeStyle.Font.IsBold = true;

            // Cells that will use the theme style.
            string[] themedCells = { "A1", "B2", "C3", "D4" };
            foreach (string address in themedCells)
            {
                Cell cell = sheet.Cells[address];
                cell.PutValue($"Themed {address}");
                cell.SetStyle(themeStyle);
            }

            // -------------------------------------------------
            // 2. Record the cells that are expected to be affected by the theme.
            // -------------------------------------------------
            List<string> cellsToWatch = new List<string>(themedCells);

            // -------------------------------------------------
            // 3. Apply a custom theme (different colors).
            // -------------------------------------------------
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(80, 80, 80),    // Text2
                Color.FromArgb(255, 0, 0),     // Accent1 (red)
                Color.FromArgb(0, 255, 0),     // Accent2 (green)
                Color.FromArgb(0, 0, 255),     // Accent3 (blue)
                Color.FromArgb(255, 255, 0),   // Accent4 (yellow)
                Color.FromArgb(255, 0, 255),   // Accent5 (magenta)
                Color.FromArgb(0, 255, 255),   // Accent6 (cyan)
                Color.FromArgb(0, 0, 128),     // Hyperlink (navy)
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink (purple)
            };

            // Apply the custom theme to the workbook.
            workbook.CustomTheme("CustomRedGreenTheme", customColors);

            // -------------------------------------------------
            // 4. After the theme change, verify each watched cell.
            // -------------------------------------------------
            Console.WriteLine("=== Theme Change Verification Log ===");
            foreach (string address in cellsToWatch)
            {
                Cell cell = sheet.Cells[address];
                Style cellStyle = cell.GetStyle();

                // If the cell's font still has a ThemeColor reference, we assume it will reflect the new theme.
                // If ThemeColor is null, the cell did not retain the theme reference and thus failed to update.
                if (cellStyle.Font.ThemeColor != null)
                {
                    // The cell is still linked to a theme color; log success.
                    Console.WriteLine($"[SUCCESS] Cell {address} retains ThemeColor ({cellStyle.Font.ThemeColor.ColorType}).");
                }
                else
                {
                    // The cell lost its theme linkage; log failure.
                    Console.WriteLine($"[FAILURE] Cell {address} does NOT retain ThemeColor after theme change.");
                }
            }

            // -------------------------------------------------
            // 5. Save the workbook (optional, just to demonstrate lifecycle compliance).
            // -------------------------------------------------
            workbook.Save("ThemeChangeLogDemo.xlsx");
        }
    }

    // Entry point for execution.
    class Program
    {
        static void Main(string[] args)
        {
            ThemeChangeLogger.Run();
        }
    }
}
