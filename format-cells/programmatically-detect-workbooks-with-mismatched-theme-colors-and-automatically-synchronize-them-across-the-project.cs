// Title: Batch synchronize mismatched Excel theme colors across workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a reference .xlsx workbook with Aspose.Cells, iterates over all other .xlsx files in a folder, compares each ThemeColorType (Accent1‑Accent6, Hyperlink, etc.) to the reference, updates any differing colors using SetThemeColor, and saves the modified files. | Generate a method that takes a source workbook path and a collection of target workbook paths, copies the full theme palette from the source to each target with GetThemeColor and SetThemeColor, and returns a list of files that were changed. | Create a console application that logs which workbooks required theme synchronization versus those already matching, using Aspose.Cells theme APIs and handling missing or corrupted files gracefully.
// Common Searches: aspnet batch update Excel theme colors with Aspose.Cells | c# compare workbook theme palette and copy missing colors | how to ensure consistent theme colors across multiple .xlsx files using Aspose.Cells | automate Excel theme synchronization for a folder of workbooks in .NET | detect theme color differences between Excel files programmatically
// Tags: Aspose.Cells batch theme palette update | C# GetThemeColor SetThemeColor usage | Excel workbook theme consistency check | Automated Excel theme color copy | Programmatic Excel theme alignment across multiple files

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For ThemeColorType enum

// The example scans a directory for .xlsx files, treats the first workbook as the reference theme source, compares each defined ThemeColorType (Accent1‑Accent6, Hyperlink, etc.) with the reference, updates any mismatched colors using SetThemeColor, overwrites the original files, and logs the synchronization results.
class ThemeSynchronizer
{
    // Define which theme colors to compare/synchronize.
    private static readonly ThemeColorType[] ThemeColorsToCheck = new ThemeColorType[]
    {
        ThemeColorType.Accent1,
        ThemeColorType.Accent2,
        ThemeColorType.Accent3,
        ThemeColorType.Accent4,
        ThemeColorType.Accent5,
        ThemeColorType.Accent6,
        ThemeColorType.Hyperlink,
        ThemeColorType.FollowedHyperlink,
        ThemeColorType.Text1,
        ThemeColorType.Background1,
        ThemeColorType.Text2,
        ThemeColorType.Background2
    };

    static void Main()
    {
        try
        {
            // Folder containing the workbooks to process.
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Load all Excel files in the folder.
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (files.Length == 0)
            {
                Console.WriteLine("No workbooks found.");
                return;
            }

            // Load the first workbook and treat its theme as the reference.
            if (!File.Exists(files[0]))
            {
                Console.WriteLine($"Reference file not found: {files[0]}");
                return;
            }

            Workbook referenceWb = new Workbook(files[0]);

            // Process remaining workbooks.
            for (int i = 1; i < files.Length; i++)
            {
                string file = files[i];

                if (!File.Exists(file))
                {
                    Console.WriteLine($"File not found, skipping: {file}");
                    continue;
                }

                try
                {
                    Workbook wb = new Workbook(file);
                    bool mismatched = false;

                    // Compare each defined theme color.
                    foreach (ThemeColorType type in ThemeColorsToCheck)
                    {
                        Color refColor = referenceWb.GetThemeColor(type);
                        Color curColor = wb.GetThemeColor(type);

                        if (!refColor.Equals(curColor))
                        {
                            mismatched = true;
                            // Synchronize the color to the reference.
                            wb.SetThemeColor(type, refColor);
                        }
                    }

                    if (mismatched)
                    {
                        // Save the workbook with synchronized theme (overwrite original).
                        wb.Save(file);
                        Console.WriteLine($"Synchronized theme for: {Path.GetFileName(file)}");
                    }
                    else
                    {
                        Console.WriteLine($"Theme already matches for: {Path.GetFileName(file)}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(file)}': {ex.Message}");
                }
            }

            Console.WriteLine("Theme synchronization completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
