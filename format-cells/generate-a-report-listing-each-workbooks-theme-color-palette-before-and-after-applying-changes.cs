// Title: Generate an Excel report of each workbook’s theme color palette before and after changes using C# and Aspose.Cells
// Common Searches: how to export Excel theme colors to a report with Aspose.Cells in C# | C# code to read and compare workbook theme palette using Aspose.Cells | batch process multiple .xlsx files to capture theme colors before and after modification | using reflection to access ThemeColor enum in Aspose.Cells for theme color extraction | create a summary workbook that lists ARGB values of theme colors pre and post update
// Tags: Aspose.Cells theme palette extraction C# | compare workbook theme colors Aspose.Cells | batch Excel theme color report Aspose.Cells | reflection ThemeColor GetColor SetColor Aspose.Cells | export theme palette ARGB values C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace ThemePaletteReportGenerator
{
    // The C# program scans a folder for .xlsx files, loads each workbook with Aspose.Cells, retrieves its theme color palette via reflection, applies sample changes to Accent1 and Accent2, captures the palette again, and writes the workbook name, theme color name, and ARGB values before and after the changes into a new Excel report saved to a specified path.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Folder containing the workbooks to process
                string inputFolder = @"C:\InputWorkbooks";
                // Path for the generated report workbook
                string reportPath = @"C:\ThemePaletteReport.xlsx";

                // Validate input folder
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Ensure the directory for the report exists
                string reportDir = Path.GetDirectoryName(reportPath);
                if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                // Create a new workbook for the report
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Name = "ThemePaletteReport";

                // Write header row
                int row = 0;
                reportSheet.Cells[row, 0].PutValue("Workbook");
                reportSheet.Cells[row, 1].PutValue("Theme Color");
                reportSheet.Cells[row, 2].PutValue("Before (ARGB)");
                reportSheet.Cells[row, 3].PutValue("After (ARGB)");
                row++;

                // Process each .xlsx file in the input folder
                foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook wb = new Workbook(filePath);

                        // Capture the theme palette before changes
                        Dictionary<string, Color> beforePalette = GetThemePalette(wb);

                        // Apply sample changes to the theme (e.g., modify Accent1 and Accent2)
                        ApplyThemeChanges(wb);

                        // Capture the theme palette after changes
                        Dictionary<string, Color> afterPalette = GetThemePalette(wb);

                        // Write the palette information to the report sheet
                        foreach (var kvp in beforePalette)
                        {
                            string themeColorName = kvp.Key;
                            reportSheet.Cells[row, 0].PutValue(Path.GetFileName(filePath));
                            reportSheet.Cells[row, 1].PutValue(themeColorName);
                            reportSheet.Cells[row, 2].PutValue(ColorToArgbString(kvp.Value));

                            // After palette may be missing a key if the theme API is unavailable; handle gracefully
                            if (afterPalette.TryGetValue(themeColorName, out Color afterColor))
                                reportSheet.Cells[row, 3].PutValue(ColorToArgbString(afterColor));
                            else
                                reportSheet.Cells[row, 3].PutValue("N/A");

                            row++;
                        }
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                    }
                }

                // Save the report workbook
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Retrieves the theme color palette of a workbook as a dictionary (key = theme color name)
        private static Dictionary<string, Color> GetThemePalette(Workbook wb)
        {
            var palette = new Dictionary<string, Color>();

            // Use reflection to avoid compile‑time dependency on Theme classes (may not exist in older versions)
            PropertyInfo themeProp = wb.GetType().GetProperty("Theme");
            if (themeProp == null) return palette;

            object theme = themeProp.GetValue(wb);
            PropertyInfo colorSchemeProp = theme?.GetType().GetProperty("ColorScheme");
            object colorScheme = colorSchemeProp?.GetValue(theme);
            if (colorScheme == null) return palette;

            // Obtain the ThemeColor enum type via reflection
            Type themeColorEnum = colorScheme.GetType().Assembly.GetType("Aspose.Cells.Drawing.ThemeColor");
            if (themeColorEnum == null) return palette;

            MethodInfo getColorMethod = colorScheme.GetType().GetMethod("GetColor", new[] { themeColorEnum });
            if (getColorMethod == null) return palette;

            foreach (object enumValue in Enum.GetValues(themeColorEnum))
            {
                object colorObj = getColorMethod.Invoke(colorScheme, new[] { enumValue });
                if (colorObj is Color color)
                {
                    palette[enumValue.ToString()] = color;
                }
            }

            return palette;
        }

        // Applies sample modifications to the workbook's theme palette using reflection
        private static void ApplyThemeChanges(Workbook wb)
        {
            PropertyInfo themeProp = wb.GetType().GetProperty("Theme");
            if (themeProp == null) return;

            object theme = themeProp.GetValue(wb);
            PropertyInfo colorSchemeProp = theme?.GetType().GetProperty("ColorScheme");
            object colorScheme = colorSchemeProp?.GetValue(theme);
            if (colorScheme == null) return;

            Type themeColorEnum = colorScheme.GetType().Assembly.GetType("Aspose.Cells.Drawing.ThemeColor");
            if (themeColorEnum == null) return;

            MethodInfo setColorMethod = colorScheme.GetType().GetMethod("SetColor", new[] { themeColorEnum, typeof(Color) });
            if (setColorMethod == null) return;

            // Example modifications: set Accent1 to Red and Accent2 to Green (if those enum members exist)
            object accent1 = Enum.Parse(themeColorEnum, "Accent1", ignoreCase: true);
            object accent2 = Enum.Parse(themeColorEnum, "Accent2", ignoreCase: true);

            setColorMethod.Invoke(colorScheme, new object[] { accent1, Color.Red });
            setColorMethod.Invoke(colorScheme, new object[] { accent2, Color.Green });
        }

        // Converts a System.Drawing.Color to an ARGB string representation
        private static string ColorToArgbString(Color color)
        {
            return $"0x{color.ToArgb():X8}";
        }
    }
}
