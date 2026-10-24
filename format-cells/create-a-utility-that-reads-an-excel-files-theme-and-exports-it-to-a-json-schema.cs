// Title: C# example: Export Excel workbook theme colors and fonts to a JSON file with Aspose.Cells
// Common Searches: C# Aspose.Cells extract workbook theme colors and save as JSON | How to read Excel theme font names using Aspose.Cells in .NET | Export Excel theme (color scheme and fonts) to JSON with Aspose.Cells | Aspose.Cells .NET retrieve ThemeColorSchemeIndex values programmatically | Convert Excel theme colors to hex strings in C# using Aspose.Cells
// Tags: aspocells theme color scheme extraction | aspocells theme font scheme retrieval | c# serialize excel theme to json | convert system.drawing.color to hex c# | reflection based access to aspocells theme api

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Retained for potential future use

namespace ExcelThemeExporter
{
    // Represents the extracted theme information.
    // The ThemeExporter class loads an Excel workbook with Aspose.Cells, uses reflection to obtain the workbook's Theme object, extracts each ThemeColorSchemeIndex color and converts it to a hex string, reads the major and minor font names from the ThemeFontScheme, populates a ThemeInfo model, serializes this model to indented JSON, and writes the result to a specified file while handling missing files and other exceptions.
    public class ThemeInfo
    {
        public Dictionary<string, string> Colors { get; set; } = new Dictionary<string, string>();
        public FontInfo Fonts { get; set; } = new FontInfo();
    }

    // Holds major and minor font names.
    public class FontInfo
    {
        public string Major { get; set; } = string.Empty;
        public string Minor { get; set; } = string.Empty;
    }

    public static class ThemeExporter
    {
        /// <param name="excelFilePath">Path to the source Excel workbook.</param>
        /// <param name="jsonOutputPath">Path where the JSON schema will be saved.</param>
        public static void ExportThemeToJson(string excelFilePath, string jsonOutputPath)
        {
            try
            {
                // Verify source file exists.
                if (!File.Exists(excelFilePath))
                    throw new FileNotFoundException($"Excel file not found: {excelFilePath}");

                // Load the workbook.
                Workbook workbook = new Workbook(excelFilePath);

                ThemeInfo themeInfo = new ThemeInfo();

                // Use reflection to access theme-related types (compatible with multiple Aspose.Cells versions).
                object themeObj = workbook.GetType().GetProperty("Theme")?.GetValue(workbook);
                if (themeObj != null)
                {
                    // ----- Extract color scheme -----
                    PropertyInfo colorSchemeProp = themeObj.GetType().GetProperty("ThemeColorScheme");
                    object colorSchemeObj = colorSchemeProp?.GetValue(themeObj);
                    if (colorSchemeObj != null)
                    {
                        // Get the enum type for ThemeColorSchemeIndex.
                        Type indexEnumType = colorSchemeObj.GetType().Assembly.GetType("Aspose.Cells.Drawing.ThemeColorSchemeIndex");
                        if (indexEnumType != null)
                        {
                            MethodInfo getColorMethod = colorSchemeObj.GetType().GetMethod("GetColor", new[] { indexEnumType });
                            foreach (object enumValue in Enum.GetValues(indexEnumType))
                            {
                                Color color = (Color)getColorMethod.Invoke(colorSchemeObj, new[] { enumValue });
                                themeInfo.Colors[enumValue.ToString()] = ColorToHex(color);
                            }
                        }
                    }

                    // ----- Extract font scheme -----
                    PropertyInfo fontSchemeProp = themeObj.GetType().GetProperty("ThemeFontScheme");
                    object fontSchemeObj = fontSchemeProp?.GetValue(themeObj);
                    if (fontSchemeObj != null)
                    {
                        PropertyInfo majorFontProp = fontSchemeObj.GetType().GetProperty("MajorFont");
                        PropertyInfo minorFontProp = fontSchemeObj.GetType().GetProperty("MinorFont");

                        object majorFont = majorFontProp?.GetValue(fontSchemeObj);
                        object minorFont = minorFontProp?.GetValue(fontSchemeObj);

                        PropertyInfo nameProp = typeof(Font).GetProperty("Name");

                        themeInfo.Fonts.Major = majorFont != null ? nameProp.GetValue(majorFont)?.ToString() ?? string.Empty : string.Empty;
                        themeInfo.Fonts.Minor = minorFont != null ? nameProp.GetValue(minorFont)?.ToString() ?? string.Empty : string.Empty;
                    }
                }

                // Serialize to JSON with indentation.
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(themeInfo, jsonOptions);

                // Ensure output directory exists.
                string outputDir = Path.GetDirectoryName(jsonOutputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Write JSON to file.
                File.WriteAllText(jsonOutputPath, json);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error exporting theme: {ex.Message}");
                throw;
            }
        }

        // Helper: converts a System.Drawing.Color to a hex string (including alpha).
        private static string ColorToHex(Color color)
        {
            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }

    // Example usage.
    class Program
    {
        static void Main(string[] args)
        {
            // Adjust paths as needed.
            string excelPath = @"C:\Data\SampleWorkbook.xlsx";
            string jsonPath = @"C:\Data\WorkbookTheme.json";

            try
            {
                ThemeExporter.ExportThemeToJson(excelPath, jsonPath);
                Console.WriteLine("Theme exported successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to export theme: {ex.Message}");
            }
        }
    }
}
