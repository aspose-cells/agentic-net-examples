// Title: Generate a before‑and‑after theme color palette report for multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read workbook theme colors and export to hex strings | How to change the first theme color in an Excel file with Aspose.Cells .NET | Create a comparison report of original and modified theme palettes for several .xlsx files | Aspose.Cells SetThemeColor example for updating theme palette programmatically | Generate Excel report listing workbook name with before and after theme palette values
// Tags: Aspose.Cells GetThemeColor API | Aspose.Cells SetThemeColor API | Excel theme palette extraction C# | Modify Excel theme colors Aspose.Cells | Theme palette comparison report Aspose.Cells | Save modified workbook with suffix Aspose.Cells | Convert Color array to hex string C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Retained for ThemeColorType enum

namespace ThemePaletteReport
{
    // The C# program loads each Excel workbook, retrieves its 12‑color theme palette using Aspose.Cells GetThemeColor, clones and modifies the first theme color to teal with SetThemeColor, saves the altered workbook with a "_Modified" suffix, and writes a separate report workbook that lists the file name together with the original and updated palette values formatted as comma‑separated hex strings.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths of workbooks to process
            string[] workbookPaths = new string[]
            {
                @"C:\Workbooks\Book1.xlsx",
                @"C:\Workbooks\Book2.xlsx"
                // add more paths as needed
            };

            // Create a new workbook that will hold the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Theme Palette Report";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Workbook");
            reportSheet.Cells[0, 1].PutValue("Before Palette (Hex)");
            reportSheet.Cells[0, 2].PutValue("After Palette (Hex)");

            int reportRow = 1; // start after header

            foreach (string path in workbookPaths)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        continue;
                    }

                    // Load the workbook
                    Workbook wb = new Workbook(path);

                    // Get the original theme color palette (array of 12 colors)
                    Color[] beforePalette = GetThemePalette(wb);

                    // ---- Apply changes to the theme palette ----
                    // Example change: set the first theme color to a custom teal color
                    Color[] afterPalette = (Color[])beforePalette.Clone();
                    afterPalette[0] = Color.FromArgb(0, 128, 128); // teal
                    SetThemePalette(wb, afterPalette);
                    // -------------------------------------------

                    // Save the modified workbook (optional)
                    string dir = Path.GetDirectoryName(path);
                    if (dir == null)
                    {
                        Console.WriteLine($"Unable to determine directory for: {path}");
                        continue;
                    }

                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(path);
                    string ext = Path.GetExtension(path);
                    string modifiedPath = Path.Combine(dir, $"{fileNameWithoutExt}_Modified{ext}");
                    wb.Save(modifiedPath);

                    // Convert palettes to comma‑separated hex strings for reporting
                    string beforeHex = PaletteToHexString(beforePalette);
                    string afterHex = PaletteToHexString(afterPalette);

                    // Write data to the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(Path.GetFileName(path));
                    reportSheet.Cells[reportRow, 1].PutValue(beforeHex);
                    reportSheet.Cells[reportRow, 2].PutValue(afterHex);
                    reportRow++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Ensure the report directory exists
            string reportPath = @"C:\Reports\ThemePaletteReport.xlsx";
            try
            {
                string? reportDir = Path.GetDirectoryName(reportPath);
                if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                // Save the report workbook
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save report: {ex.Message}");
            }
        }

        // Retrieves the theme color palette (12 colors) from a workbook
        private static Color[] GetThemePalette(Workbook wb)
        {
            Color[] palette = new Color[12];
            for (int i = 0; i < 12; i++)
            {
                // Use Workbook's GetThemeColor method (available in all supported versions)
                palette[i] = wb.GetThemeColor((ThemeColorType)i);
            }
            return palette;
        }

        // Sets a new theme color palette for a workbook
        private static void SetThemePalette(Workbook wb, Color[] palette)
        {
            if (palette == null || palette.Length != 12)
                throw new ArgumentException("Palette must contain exactly 12 colors.");

            for (int i = 0; i < 12; i++)
            {
                // Use Workbook's SetThemeColor method
                wb.SetThemeColor((ThemeColorType)i, palette[i]);
            }
        }

        // Converts an array of Colors to a single comma‑separated hex string (e.g., "#FF0000")
        private static string PaletteToHexString(Color[] palette)
        {
            List<string> hexValues = new List<string>();
            foreach (Color c in palette)
            {
                // Ensure ARGB is represented as #RRGGBB (ignore alpha for theme colors)
                string hex = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                hexValues.Add(hex);
            }
            return string.Join(", ", hexValues);
        }
    }
}
