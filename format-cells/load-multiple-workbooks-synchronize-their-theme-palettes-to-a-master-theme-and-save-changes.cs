// Title: Batch Synchronize Excel Workbook Themes with Aspose.Cells for .NET
// Description: Loads a master workbook, iterates through a list of Excel files, copies the master’s theme to each using Aspose.Cells CopyTheme, and saves the changes while gracefully handling missing files.
// Keywords: Aspose.Cells | CopyTheme | C# | .NET | Excel theme synchronization | batch theme update | master workbook theme | apply Excel theme programmatically | theme palette copy
// Common Searches: copy Excel theme from one workbook to another Aspose.Cells | batch update Excel theme C# Aspose.Cells | synchronize workbook themes programmatically .NET | apply corporate theme to multiple spreadsheets using Aspose.Cells | automate Excel theme copying C#
// Developer Intent: Copy a master workbook’s theme to several other Excel files and persist the updates.
// Use Cases: Enforce corporate branding by applying a single theme to all department spreadsheets in a shared folder. | Refresh legacy reports after a style redesign by batch‑updating their color palettes. | Integrate theme synchronization into a CI/CD pipeline to guarantee consistent visual appearance of generated Excel outputs.
// AI Prompts: Generate C# code that accepts a master workbook path and an array of target workbook paths, uses Aspose.Cells to copy the master theme to each existing target, skips missing files, and returns a processing summary. | Create a reusable Aspose.Cells method that takes a source Workbook object and a collection of destination Workbook objects, applies CopyTheme to each, logs success or error messages, and saves the results.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeSync
{
    // Loads a master workbook, iterates through a list of Excel files, copies the master’s theme to each using Aspose.Cells CopyTheme, and saves the changes while gracefully handling missing files.
    public class ThemeSynchronizer
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Paths of the workbooks to be synchronized.
            // The first workbook in the list is considered the master theme source.
            string[] workbookPaths = new string[]
            {
                "MasterWorkbook.xlsx",
                "WorkbookA.xlsx",
                "WorkbookB.xlsx",
                "WorkbookC.xlsx"
            };

            // Verify that the master workbook exists.
            if (!File.Exists(workbookPaths[0]))
            {
                Console.WriteLine($"Master workbook not found: {workbookPaths[0]}");
                return;
            }

            // Load the master workbook whose theme will be copied to the others.
            using (Workbook masterWorkbook = new Workbook(workbookPaths[0]))
            {
                // Iterate over the remaining workbooks, copy the master theme, and save them.
                for (int i = 1; i < workbookPaths.Length; i++)
                {
                    string targetPath = workbookPaths[i];

                    // Verify that the target workbook exists before processing.
                    if (!File.Exists(targetPath))
                    {
                        Console.WriteLine($"Target workbook not found, skipping: {targetPath}");
                        continue;
                    }

                    // Load the target workbook.
                    using (Workbook targetWorkbook = new Workbook(targetPath))
                    {
                        // Copy the theme from the master workbook to the target workbook.
                        targetWorkbook.CopyTheme(masterWorkbook);

                        // Save the updated workbook (overwrites the original file).
                        targetWorkbook.Save(targetPath);
                        Console.WriteLine($"Theme synchronized and saved: {targetPath}");
                    }
                }
            }
        }
    }
}
