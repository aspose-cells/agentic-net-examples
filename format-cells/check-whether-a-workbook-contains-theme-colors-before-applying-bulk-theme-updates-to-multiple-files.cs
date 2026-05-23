using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example file list – replace with actual file paths
            string[] workbookFiles = new string[]
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
            };

            ApplyBulkThemeUpdates(workbookFiles);
        }

        /// <summary>
        /// Loads each workbook, checks if it has a theme, and if so applies bulk theme color updates.
        /// </summary>
        /// <param name="files">Array of workbook file paths.</param>
        static void ApplyBulkThemeUpdates(string[] files)
        {
            foreach (string filePath in files)
            {
                // Load the workbook (uses the provided load rule)
                Workbook workbook = new Workbook(filePath);

                // Check whether the workbook contains a theme.
                // The Theme property returns the theme name; an empty string means no theme.
                if (!string.IsNullOrEmpty(workbook.Theme))
                {
                    // Example bulk update: change a few theme colors.
                    // You can modify any ThemeColorType as needed.
                    workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 0, 120, 215)); // Blue accent
                    workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 232, 17, 35)); // Red accent
                    workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(255, 255, 185, 0)); // Yellow accent

                    // Save the workbook (uses the provided save rule)
                    workbook.Save(filePath);
                    Console.WriteLine($"Theme updated and saved for: {filePath}");
                }
                else
                {
                    Console.WriteLine($"No theme found in workbook: {filePath}. Skipping update.");
                }

                // Dispose the workbook to release resources.
                workbook.Dispose();
            }
        }
    }
}