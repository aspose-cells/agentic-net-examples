// Title: Programmatically synchronize Excel workbook theme colors across multiple files using Aspose.Cells for .NET
// Common Searches: how to copy theme colors from a reference Excel workbook to other workbooks with Aspose.Cells C# | batch update mismatched theme colors in multiple .xlsx files using Aspose.Cells .NET | detect and align Excel workbook theme colors across a project in C# | Aspose.Cells set specific ThemeColorType programmatically in .NET | C# code to ensure consistent Excel themes in a set of workbooks
// Tags: Aspose.Cells set theme color programmatically | synchronize Excel workbook themes batch | detect mismatched Excel theme colors .NET | copy theme colors between workbooks Aspose | align Excel theme across multiple files C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeSyncDemo
{
    // The example defines a ThemeSynchronizer class that loads a reference workbook, extracts its theme colors via GetThemeColor, and then iterates through additional workbooks. For each workbook it compares every ThemeColorType with the reference, applies any differing colors using SetThemeColor, and saves the file only when changes are needed, handling missing files and exceptions while logging the process.
    public static class ThemeSynchronizer
    {
        /// <param name="workbookPaths">Full paths of the workbooks to process.</param>
        public static void SynchronizeThemes(IEnumerable<string> workbookPaths)
        {
            if (workbookPaths == null) throw new ArgumentNullException(nameof(workbookPaths));

            var paths = new List<string>(workbookPaths);
            if (paths.Count == 0) return; // Nothing to do

            // Ensure the reference workbook exists
            string referencePath = paths[0];
            if (!File.Exists(referencePath))
            {
                Console.WriteLine($"Reference file not found: {referencePath}");
                return;
            }

            // Load the reference workbook and capture its theme colors
            Dictionary<ThemeColorType, Color> referenceColors;
            try
            {
                var referenceWorkbook = new Workbook(referencePath);
                referenceColors = GetThemeColors(referenceWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load reference workbook '{referencePath}': {ex.Message}");
                return;
            }

            // Process remaining workbooks
            for (int i = 1; i < paths.Count; i++)
            {
                string path = paths[i];
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                try
                {
                    var wb = new Workbook(path);
                    var currentColors = GetThemeColors(wb);

                    bool needsSave = false;
                    foreach (var kvp in referenceColors)
                    {
                        ThemeColorType type = kvp.Key;
                        Color refColor = kvp.Value;
                        Color curColor = currentColors.ContainsKey(type) ? currentColors[type] : Color.Empty;

                        if (curColor.ToArgb() != refColor.ToArgb())
                        {
                            // Apply the reference color to the workbook
                            wb.SetThemeColor(type, refColor);
                            needsSave = true;
                        }
                    }

                    if (needsSave)
                    {
                        try
                        {
                            wb.Save(path, SaveFormat.Xlsx);
                            Console.WriteLine($"Synchronized theme for: {path}");
                        }
                        catch (Exception saveEx)
                        {
                            Console.WriteLine($"Failed to save workbook '{path}': {saveEx.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Theme already synchronized: {path}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Retrieves all theme colors defined in a workbook.
        /// </summary>
        /// <param name="workbook">The workbook to inspect.</param>
        /// <returns>Dictionary of theme colors.</returns>
        private static Dictionary<ThemeColorType, Color> GetThemeColors(Workbook workbook)
        {
            var dict = new Dictionary<ThemeColorType, Color>();

            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                try
                {
                    // Use Workbook.GetThemeColor to obtain the color for the given type
                    Color clr = workbook.GetThemeColor(type);
                    dict[type] = clr;
                }
                catch
                {
                    // Ignore unsupported types
                }
            }

            return dict;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            var workbookFiles = new List<string>
            {
                @"C:\Projects\Workbook1.xlsx",
                @"C:\Projects\Workbook2.xlsx",
                @"C:\Projects\Workbook3.xlsx"
            };

            ThemeSynchronizer.SynchronizeThemes(workbookFiles);
        }
    }
}
