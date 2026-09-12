// Title: Report original and modified Excel theme color palettes for multiple workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads each .xlsx file, extracts the workbook's theme colors via reflection, prints the original palette, changes Accent1 to dark green and Hyperlink to blue, prints the updated palette, and saves the workbook with a '_Modified' suffix. | Add functionality to export both the original and the modified theme palettes to a CSV file named after each processed workbook. | Create a method that converts the theme palette dictionary into a JSON string for integration with web APIs.
// Common Searches: how to list Excel theme colors using Aspose.Cells C# reflection | Aspose.Cells change Accent1 theme color programmatically | generate before and after theme palette report for multiple .xlsx files in .NET | save modified workbook after updating theme colors with Aspose.Cells | export Excel theme palette to CSV using Aspose.Cells C#
// Tags: aspocells retrieve theme colors via reflection | aspocells modify theme palette accent1 dark green | aspocells export theme palette to csv | aspocells save workbook with updated theme colors | c# generate before after theme palette report

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for ThemeColorType

// The example iterates over a list of .xlsx files, uses reflection to access each workbook's Theme object, prints the original theme color palette, updates Accent1 and Hyperlink colors, prints the modified palette, and saves the workbook with a '_Modified' suffix. Optional extensions show how to export palettes to CSV or JSON.
class ThemePaletteReport
{
    // Retrieves all theme colors from a workbook as a dictionary using reflection
    static Dictionary<string, Color> GetThemeColors(Workbook wb)
    {
        var colors = new Dictionary<string, Color>();

        try
        {
            // Obtain the Theme object via reflection
            object themeObj = wb.GetType().GetProperty("Theme")?.GetValue(wb);
            if (themeObj == null)
                return colors;

            MethodInfo getMethod = themeObj.GetType().GetMethod(
                "GetThemeColor", new[] { typeof(ThemeColorType) });

            if (getMethod == null)
                return colors;

            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Invoke GetThemeColor for each enum value
                Color color = (Color)getMethod.Invoke(themeObj, new object[] { type });
                colors[type.ToString()] = color;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving theme colors: {ex.Message}");
        }

        return colors;
    }

    // Prints the theme palette to the console
    static void PrintThemePalette(string title, Dictionary<string, Color> palette)
    {
        Console.WriteLine(title);
        foreach (var kvp in palette)
        {
            Console.WriteLine($"{kvp.Key}: #{kvp.Value.ToArgb():X8}");
        }
        Console.WriteLine();
    }

    static void Main()
    {
        // List of workbook file paths to process
        string[] workbookFiles = { "Workbook1.xlsx", "Workbook2.xlsx" };

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Verify that the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook wb = new Workbook(filePath);

                // Display original theme palette
                var originalPalette = GetThemeColors(wb);
                PrintThemePalette($"Original Theme Palette for '{filePath}':", originalPalette);

                // ----- Apply changes to the theme palette using reflection -----
                object themeObj = wb.GetType().GetProperty("Theme")?.GetValue(wb);
                if (themeObj != null)
                {
                    MethodInfo setMethod = themeObj.GetType().GetMethod(
                        "SetThemeColor", new[] { typeof(ThemeColorType), typeof(Color) });

                    if (setMethod != null)
                    {
                        setMethod.Invoke(themeObj, new object[]
                        {
                            ThemeColorType.Accent1,
                            Color.FromArgb(0xFF, 0x00, 0x80, 0x00) // Dark Green
                        });

                        setMethod.Invoke(themeObj, new object[]
                        {
                            ThemeColorType.Hyperlink,
                            Color.FromArgb(0xFF, 0x00, 0x00, 0xFF) // Blue
                        });
                    }
                }

                // Display modified theme palette
                var modifiedPalette = GetThemeColors(wb);
                PrintThemePalette($"Modified Theme Palette for '{filePath}':", modifiedPalette);

                // Prepare output path
                string outputDirectory = Path.GetDirectoryName(Path.GetFullPath(filePath));
                if (string.IsNullOrEmpty(outputDirectory))
                {
                    outputDirectory = Directory.GetCurrentDirectory();
                }

                string outputPath = Path.Combine(
                    outputDirectory,
                    Path.GetFileNameWithoutExtension(filePath) + "_Modified.xlsx");

                // Ensure output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Save the modified workbook
                wb.Save(outputPath);
                Console.WriteLine($"Modified workbook saved to: {outputPath}\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }
    }
}
