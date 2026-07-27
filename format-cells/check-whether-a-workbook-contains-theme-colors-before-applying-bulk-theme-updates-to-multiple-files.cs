using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeCheckDemo
{
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            string[] workbookFiles = new string[]
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
            };

            // Define a new theme color to apply (example: change Accent1 to Orange)
            ThemeColorType targetThemeColor = ThemeColorType.Accent1;
            Color newColor = Color.Orange;

            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(filePath);

                // Check if the workbook has a theme name (default theme is always present,
                // but a missing or empty name indicates no theme information)
                if (!string.IsNullOrEmpty(workbook.Theme))
                {
                    // The workbook contains a theme, so apply the bulk update
                    workbook.SetThemeColor(targetThemeColor, newColor);
                }
                else
                {
                    // No theme detected; you may choose to skip or apply a custom theme here
                    Console.WriteLine($"Workbook '{filePath}' does not contain a theme. Skipping update.");
                }

                // Save the workbook back to the same file (lifecycle rule: save)
                workbook.Save(filePath);
            }

            Console.WriteLine("Theme update process completed.");
        }
    }
}