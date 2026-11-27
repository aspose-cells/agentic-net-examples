// Title: Export an Excel workbook's theme (color scheme and font settings) to JSON with Aspose.Cells for .NET (C#)
// Common Searches: asp.net core export excel theme colors as json using aspose.cells | c# read excel workbook theme font settings with aspose.cells | how to extract theme color scheme from an xlsx file in c# | serialize excel theme information to json file using aspose.cells library | use reflection to access Aspose.Cells Theme object in c# example
// Tags: Aspose.Cells theme extraction to JSON | C# retrieve Excel theme colors | C# retrieve Excel theme fonts | Reflection usage for Aspose.Cells theme API | Generate JSON schema from workbook theme

using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Text.Json;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Retained for potential future use

namespace ExcelThemeExporter
{
    // Represents the JSON schema for an Excel theme.
    // The ThemeExporter class validates the input Excel file, loads it with Aspose.Cells, and uses reflection to access the workbook's Theme object. It extracts the ThemeColorScheme by invoking GetColor for each enum value, converts colors to hex strings, and stores them in a Colors dictionary. It also extracts the ThemeFontScheme by invoking methods to obtain font name, size, bold, italic, and color for each enum value, populating a Fonts dictionary of FontInfo objects. The collected ThemeSchema is then serialized to an indented JSON file at the specified output path, with comprehensive error handling for file I/O and reflection failures.
    public class ThemeSchema
    {
        public Dictionary<string, string> Colors { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, FontInfo> Fonts { get; set; } = new Dictionary<string, FontInfo>();
    }

    // Simple DTO for font information.
    public class FontInfo
    {
        public string? Name { get; set; }
        public double Size { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string? Color { get; set; }
    }

    public static class ThemeExporter
    {
        // Exports the theme of the given Excel file to a JSON file.
        public static void ExportThemeToJson(string excelPath, string jsonOutputPath)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Excel file not found: {excelPath}");

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load workbook: {ex.Message}", ex);
            }

            // Prepare schema container.
            ThemeSchema schema = new ThemeSchema();

            // Use reflection to access theme information (avoids compile‑time dependency on specific Aspose.Cells versions).
            try
            {
                object? theme = workbook.Theme;
                if (theme != null)
                {
                    // ----- Extract color scheme -----
                    PropertyInfo? colorSchemeProp = theme.GetType().GetProperty("ThemeColorScheme");
                    object? colorSchemeObj = colorSchemeProp?.GetValue(theme);
                    if (colorSchemeObj != null)
                    {
                        MethodInfo? getColorMethod = colorSchemeObj.GetType().GetMethod("GetColor");
                        ParameterInfo[]? parameters = getColorMethod?.GetParameters();
                        if (parameters?.Length == 1)
                        {
                            Type enumType = parameters[0].ParameterType;
                            foreach (object enumValue in Enum.GetValues(enumType))
                            {
                                try
                                {
                                    Color color = (Color)getColorMethod!.Invoke(colorSchemeObj, new[] { enumValue })!;
                                    schema.Colors[enumValue.ToString()] = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                                }
                                catch
                                {
                                    // Skip unsupported indexes.
                                }
                            }
                        }
                    }

                    // ----- Extract font scheme -----
                    PropertyInfo? fontSchemeProp = theme.GetType().GetProperty("ThemeFontScheme");
                    object? fontSchemeObj = fontSchemeProp?.GetValue(theme);
                    if (fontSchemeObj != null)
                    {
                        MethodInfo? getFontNameMethod = fontSchemeObj.GetType().GetMethod("GetFontName");
                        MethodInfo? getFontSizeMethod = fontSchemeObj.GetType().GetMethod("GetFontSize");
                        MethodInfo? getBoldMethod = fontSchemeObj.GetType().GetMethod("GetBold");
                        MethodInfo? getItalicMethod = fontSchemeObj.GetType().GetMethod("GetItalic");
                        MethodInfo? getFontColorMethod = fontSchemeObj.GetType().GetMethod("GetColor");

                        ParameterInfo[]? parameters = getFontNameMethod?.GetParameters();
                        if (parameters?.Length == 1)
                        {
                            Type enumType = parameters[0].ParameterType;
                            foreach (object enumValue in Enum.GetValues(enumType))
                            {
                                try
                                {
                                    FontInfo fontInfo = new FontInfo
                                    {
                                        Name = (string?)getFontNameMethod!.Invoke(fontSchemeObj, new[] { enumValue }) ?? string.Empty,
                                        Size = Convert.ToDouble(getFontSizeMethod!.Invoke(fontSchemeObj, new[] { enumValue }) ?? 0.0),
                                        Bold = (bool)(getBoldMethod!.Invoke(fontSchemeObj, new[] { enumValue }) ?? false),
                                        Italic = (bool)(getItalicMethod!.Invoke(fontSchemeObj, new[] { enumValue }) ?? false)
                                    };

                                    Color fontColor = (Color)getFontColorMethod!.Invoke(fontSchemeObj, new[] { enumValue })!;
                                    fontInfo.Color = $"#{fontColor.R:X2}{fontColor.G:X2}{fontColor.B:X2}";

                                    schema.Fonts[enumValue.ToString()] = fontInfo;
                                }
                                catch
                                {
                                    // Skip unsupported indexes.
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If reflection fails, continue with whatever data has been collected.
                Console.WriteLine($"Warning: Unable to extract full theme information. {ex.Message}");
            }

            // Serialize schema to JSON.
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(schema, jsonOptions);

            // Ensure output directory exists.
            string? outputDir = Path.GetDirectoryName(jsonOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Write JSON to file.
            try
            {
                File.WriteAllText(jsonOutputPath, json);
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to write JSON file: {ex.Message}", ex);
            }
        }
    }

    // Example usage.
    class Program
    {
        static void Main(string[] args)
        {
            // Adjust paths as needed.
            string excelFile = @"C:\Input\Sample.xlsx";
            string jsonFile = @"C:\Output\SampleTheme.json";

            try
            {
                ThemeExporter.ExportThemeToJson(excelFile, jsonFile);
                Console.WriteLine("Theme exported to JSON successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
