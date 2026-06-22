using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace ThemeSynchronizationDemo
{
    class Program
    {
        static void Main()
        {
            // List of workbook file paths in the project
            string[] workbookPaths = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
                // Add more paths as needed
            };

            if (workbookPaths.Length == 0)
            {
                Console.WriteLine("No workbooks specified.");
                return;
            }

            // Load the reference workbook (first in the list)
            Workbook referenceWorkbook = new Workbook(workbookPaths[0]);

            // Iterate over the remaining workbooks and synchronize their themes if needed
            for (int i = 1; i < workbookPaths.Length; i++)
            {
                string path = workbookPaths[i];
                Workbook targetWorkbook = new Workbook(path);

                if (!ThemesAreEqual(referenceWorkbook, targetWorkbook))
                {
                    // Copy the theme from the reference workbook to the target workbook
                    targetWorkbook.CopyTheme(referenceWorkbook);
                    // Save the updated workbook (overwrites the original file)
                    targetWorkbook.Save(path);
                    Console.WriteLine($"Theme synchronized for: {path}");
                }
                else
                {
                    Console.WriteLine($"Theme already matches for: {path}");
                }

                // Dispose the workbook to free resources
                targetWorkbook.Dispose();
            }

            // Dispose the reference workbook
            referenceWorkbook.Dispose();
        }

        // Compares theme colors of two workbooks
        private static bool ThemesAreEqual(Workbook wb1, Workbook wb2)
        {
            // Compare each ThemeColorType (excluding StyleColor which is internal)
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Skip the StyleColor enum value (value 12) as it is not a real theme color
                if (type == ThemeColorType.StyleColor) continue;

                Color color1 = wb1.GetThemeColor(type);
                Color color2 = wb2.GetThemeColor(type);

                if (!color1.Equals(color2))
                {
                    return false; // Mismatch found
                }
            }
            return true; // All theme colors match
        }
    }
}