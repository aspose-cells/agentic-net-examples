// Title: Detect Excel Workbook Theme and Apply Bulk Theme Color Changes with Aspose.Cells for .NET
// Description: Loads each workbook from a list, checks the Workbook.Theme property to see if a theme exists, and if so updates Accent1 and Accent2 colors using SetThemeColor, saves the file with an "_Updated" suffix, and releases resources.
// Keywords: Aspose.Cells C# theme detection | Workbook.Theme property | SetThemeColor example | bulk Excel theme color update | Aspose.Cells bulk processing | C# change Excel theme accents | detect Excel theme before save
// Common Searches: how to check if an Excel file has a theme using Aspose.Cells | apply bulk theme color changes to multiple workbooks C# | Aspose.Cells SetThemeColor after verifying theme presence | skip Excel files without a theme when updating colors | C# code to bulk update Excel theme accents with Aspose.Cells
// Developer Intent: Identify whether a workbook contains a theme and, only when a theme is present, modify its accent colors in bulk.
// Use Cases: Batch‑process corporate spreadsheets to enforce brand colors while leaving non‑themed files untouched. | Generate a report of workbooks that lack a theme before performing any theme modifications. | Automate migration of legacy Excel files to a new corporate theme by updating Accent1 and Accent2 across all themed workbooks.
// AI Prompts: Create a C# method that receives a file path, checks if Workbook.Theme is non‑empty, changes Accent1 and Accent2 via SetThemeColor, saves the workbook with an "_Updated" suffix, and logs the outcome. | Provide Aspose.Cells code to iterate over a list of Excel files, detect the presence of a theme, apply new RGB values to Accent1 and Accent2 in bulk, skip files without a theme, and output status messages for each file.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace ThemeBulkUpdateDemo
{
    // Loads each workbook from a list, checks the Workbook.Theme property to see if a theme exists, and if so updates Accent1 and Accent2 colors using SetThemeColor, saves the file with an "_Updated" suffix, and releases resources.
    class Program
    {
        // List of workbook file paths to process
        static readonly List<string> WorkbookFiles = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx",
            "Book3.xlsx"
        };

        static void Main()
        {
            foreach (string filePath in WorkbookFiles)
            {
                ProcessWorkbook(filePath);
            }
        }

        /// <param name="filePath">Path to the workbook file.</param>
        static void ProcessWorkbook(string filePath)
        {
            // Load the workbook (lifecycle rule: use constructor with file path)
            Workbook workbook = new Workbook(filePath);

            // Check if the workbook has a theme name (non‑empty string indicates a theme is present)
            bool hasTheme = !string.IsNullOrEmpty(workbook.Theme);
            Console.WriteLine($"Workbook: {filePath}");
            Console.WriteLine($"  Theme present: {hasTheme}");
            if (hasTheme)
            {
                // Example bulk update: change Accent1 and Accent2 colors
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 0, 120, 215)); // a shade of blue
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 232, 17, 35)); // a shade of red

                // Save the updated workbook (lifecycle rule: use Save method)
                string updatedPath = System.IO.Path.GetFileNameWithoutExtension(filePath) + "_Updated.xlsx";
                workbook.Save(updatedPath);
                Console.WriteLine($"  Updated workbook saved as: {updatedPath}");
            }
            else
            {
                Console.WriteLine("  No theme detected – skipping bulk update.");
            }

            // Dispose the workbook to release resources
            workbook.Dispose();
        }
    }
}
