// Title: Batch Synchronize Excel Theme Colors with Aspose.Cells for .NET (C#)
// Description: C# utility that scans a folder of .xlsx files, uses the first workbook as a reference theme, detects mismatched ThemeColorType values, copies the reference theme to each out‑of‑sync file, and safely skips password‑protected or corrupted workbooks.
// Keywords: Aspose.Cells C# theme synchronization | compare Excel workbook theme colors | batch update Excel theme .NET | CopyTheme Aspose.Cells example | detect mismatched theme colors | automate Excel branding | folder processing Excel files | handle password protected Excel
// Common Searches: how to copy theme from one Excel file to another using Aspose.Cells | batch change Excel theme colors with C# | detect and fix mismatched workbook themes .NET | Aspose.Cells example for theme synchronization | C# script to enforce corporate Excel theme
// Developer Intent: Find a ready‑to‑run C# sample that identifies workbooks whose theme colors differ from a reference file and applies the reference theme automatically.
// Use Cases: Enforce corporate branding by ensuring all generated reports share the same theme palette. | Refresh the visual style of legacy Excel files after a design update without manual editing. | Process large collections of workbooks on a server, skipping password‑protected or corrupted files.
// AI Prompts: Generate C# code that loads a reference workbook, compares ThemeColorType values of other .xlsx files in a directory, and uses CopyTheme to align mismatched themes with Aspose.Cells. | Create a function that returns a list of Excel files whose theme colors do not match a given reference workbook using Aspose.Cells for .NET. | Write error‑handling logic for loading workbooks that may be password‑protected, logging skipped files during batch theme synchronization.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeSynchronizationDemo
{
    // C# utility that scans a folder of .xlsx files, uses the first workbook as a reference theme, detects mismatched ThemeColorType values, copies the reference theme to each out‑of‑sync file, and safely skips password‑protected or corrupted workbooks.
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to process
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("The specified folder does not exist.");
                return;
            }

            // Get all Excel files in the folder (you can adjust the pattern as needed)
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (workbookFiles.Length == 0)
            {
                Console.WriteLine("No workbooks found in the specified folder.");
                return;
            }

            // Load the first workbook – it will serve as the reference theme source
            Workbook referenceWorkbook;
            try
            {
                if (!File.Exists(workbookFiles[0]))
                {
                    Console.WriteLine($"Reference workbook not found: {workbookFiles[0]}");
                    return;
                }

                referenceWorkbook = new Workbook(workbookFiles[0]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load reference workbook: {ex.Message}");
                return;
            }

            // Cache the reference theme colors for quick comparison
            var referenceColors = new Dictionary<ThemeColorType, Color>();
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // ThemeColorType.StyleColor is internal; skip it
                if (type == ThemeColorType.StyleColor) continue;
                referenceColors[type] = referenceWorkbook.GetThemeColor(type);
            }

            // Process remaining workbooks
            for (int i = 1; i < workbookFiles.Length; i++)
            {
                string filePath = workbookFiles[i];
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                Workbook targetWorkbook = null;
                try
                {
                    // Attempt to load the workbook; catch password‑protected or other load errors
                    targetWorkbook = new Workbook(filePath);
                }
                catch (Exception ex)
                {
                    // Simple detection of password‑protected files via message text
                    if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"Password‑protected file skipped: {Path.GetFileName(filePath)}");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to load workbook '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                    continue;
                }

                bool mismatched = false;

                // Compare each theme color with the reference
                foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
                {
                    if (type == ThemeColorType.StyleColor) continue;

                    Color targetColor = targetWorkbook.GetThemeColor(type);
                    Color refColor = referenceColors[type];

                    if (!ColorsAreEqual(targetColor, refColor))
                    {
                        mismatched = true;
                        break; // No need to check further once a mismatch is found
                    }
                }

                if (mismatched)
                {
                    try
                    {
                        // Synchronize the theme by copying from the reference workbook
                        targetWorkbook.CopyTheme(referenceWorkbook);
                        // Save the updated workbook (overwrite the original)
                        targetWorkbook.Save(filePath);
                        Console.WriteLine($"Synchronized theme for: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to synchronize theme for '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Theme already matches for: {Path.GetFileName(filePath)}");
                }

                // Dispose the workbook to free resources
                targetWorkbook.Dispose();
            }

            // Dispose the reference workbook
            referenceWorkbook.Dispose();
        }

        // Helper method to compare two Color objects (ignores alpha channel)
        private static bool ColorsAreEqual(Color c1, Color c2)
        {
            return c1.R == c2.R && c1.G == c2.G && c1.B == c2.B;
        }
    }
}
