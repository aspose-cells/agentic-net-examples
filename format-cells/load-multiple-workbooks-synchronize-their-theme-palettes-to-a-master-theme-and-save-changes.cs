using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeSync
{
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
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the master workbook whose theme will be applied to others
            string masterWorkbookPath = "MasterThemeWorkbook.xlsx";

            // Verify master workbook exists
            if (!File.Exists(masterWorkbookPath))
            {
                Console.Error.WriteLine($"Master workbook not found: {masterWorkbookPath}");
                return;
            }

            // List of workbook file paths that need to be synchronized with the master theme
            List<string> targetWorkbookPaths = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
                // Add more paths as needed
            };

            // Load the master workbook (theme source)
            using (Workbook masterWorkbook = new Workbook(masterWorkbookPath))
            {
                // Iterate over each target workbook, copy the master theme, and save the changes
                foreach (string targetPath in targetWorkbookPaths)
                {
                    try
                    {
                        // Verify target workbook exists
                        if (!File.Exists(targetPath))
                        {
                            Console.Error.WriteLine($"Target workbook not found: {targetPath}");
                            continue;
                        }

                        // Load the target workbook
                        using (Workbook targetWorkbook = new Workbook(targetPath))
                        {
                            // Copy the theme from the master workbook to the target workbook
                            targetWorkbook.CopyTheme(masterWorkbook);

                            // Save the updated workbook (overwrites the original file)
                            targetWorkbook.Save(targetPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error processing '{targetPath}': {ex.Message}");
                    }
                }
            }
        }
    }
}