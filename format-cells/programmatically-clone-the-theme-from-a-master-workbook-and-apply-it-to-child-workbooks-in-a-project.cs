using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ThemeCloner
{
    public class ThemeManager
    {
        /// <summary>
        /// Clones the theme from a master workbook and applies it to each child workbook.
        /// </summary>
        /// <param name="masterPath">Full path to the master workbook that contains the desired theme.</param>
        /// <param name="childPaths">Collection of full paths to child workbooks that will receive the theme.</param>
        public static void ApplyThemeToChildren(string masterPath, IEnumerable<string> childPaths)
        {
            // Verify master workbook exists
            if (!File.Exists(masterPath))
            {
                Console.WriteLine($"Master workbook not found: {masterPath}");
                return;
            }

            Workbook masterWorkbook = null;
            try
            {
                // Load the master workbook (theme is automatically loaded with the file)
                masterWorkbook = new Workbook(masterPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load master workbook: {ex.Message}");
                return;
            }

            // Iterate over each child workbook path
            foreach (string childPath in childPaths)
            {
                // Verify child workbook exists
                if (!File.Exists(childPath))
                {
                    Console.WriteLine($"Child workbook not found, skipping: {childPath}");
                    continue;
                }

                try
                {
                    // Load the child workbook
                    Workbook childWorkbook = new Workbook(childPath);

                    // Copy the theme from the master workbook to the child workbook
                    childWorkbook.CopyTheme(masterWorkbook);

                    // Save the updated child workbook (overwrites the original file)
                    childWorkbook.Save(childPath);

                    Console.WriteLine($"Theme applied to: {childPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{childPath}': {ex.Message}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                // Path to the master workbook that defines the theme
                string masterWorkbookPath = @"C:\Projects\MasterTemplate.xlsx";

                // List of child workbook paths that need the theme applied
                List<string> childWorkbookPaths = new List<string>
                {
                    @"C:\Projects\Report1.xlsx",
                    @"C:\Projects\Report2.xlsx",
                    @"C:\Projects\Report3.xlsx"
                };

                // Apply the master theme to all child workbooks
                ApplyThemeToChildren(masterWorkbookPath, childWorkbookPaths);

                Console.WriteLine("Theme cloning completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}