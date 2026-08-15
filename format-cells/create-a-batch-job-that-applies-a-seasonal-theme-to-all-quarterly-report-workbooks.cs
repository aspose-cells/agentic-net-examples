// Title: Batch Apply a Seasonal Theme to Multiple Excel Workbooks with Aspose.Cells for .NET
// Description: A C# utility that iterates over a list of quarterly report files, creates a custom 12‑color seasonal theme (e.g., Spring), applies it with Workbook.CustomTheme, and saves each workbook with a "_Seasonal" suffix. Ideal for automating corporate branding across Q1‑Q4 reports.
// Keywords: Aspose.Cells | C# Excel automation | custom Excel theme | batch workbook processing | seasonal color palette | Workbook.CustomTheme | quarterly report styling | Excel theming .NET | automated Excel branding | multiple file Excel theme
// Common Searches: apply a custom theme to many Excel files using Aspose.Cells C# | batch add seasonal colors to quarterly reports | Aspose.Cells example for workbook theming | automate Excel theme application across multiple workbooks | C# code to set custom Excel theme for a folder of files
// Developer Intent: Programmatically apply a predefined seasonal color theme to a collection of quarterly Excel workbooks and save themed copies automatically.
// Use Cases: Produce branded Q1‑Q4 reports with a spring or winter visual style before distribution. | Integrate seasonal theming into a nightly build that updates financial dashboards. | Create a reusable command‑line tool that adds any custom theme to a batch of Excel files.
// AI Prompts: Generate code to read theme colors from a JSON configuration file instead of hard‑coding them. | Show how to select a different seasonal palette (e.g., Winter) based on the current month. | Add robust error handling and logging for missing files, permission issues, and save failures in the batch job.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace SeasonalThemeBatchJob
{
    // A C# utility that iterates over a list of quarterly report files, creates a custom 12‑color seasonal theme (e.g., Spring), applies it with Workbook.CustomTheme, and saves each workbook with a "_Seasonal" suffix. Ideal for automating corporate branding across Q1‑Q4 reports.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the list of quarterly report workbook file paths
            List<string> inputFiles = new List<string>
            {
                @"C:\Reports\Q1_Report.xlsx",
                @"C:\Reports\Q2_Report.xlsx",
                @"C:\Reports\Q3_Report.xlsx",
                @"C:\Reports\Q4_Report.xlsx"
            };

            // Process each workbook and apply the seasonal theme
            foreach (string inputPath in inputFiles)
            {
                // Generate an output path (e.g., add "_Seasonal" suffix before extension)
                string outputPath = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(inputPath),
                    System.IO.Path.GetFileNameWithoutExtension(inputPath) + "_Seasonal" + System.IO.Path.GetExtension(inputPath));

                ApplySeasonalTheme(inputPath, outputPath);
                Console.WriteLine($"Applied seasonal theme to: {inputPath}");
                Console.WriteLine($"Saved themed workbook as: {outputPath}");
            }
        }

        /// <param name="inputPath">Full path of the source workbook.</param>
        /// <param name="outputPath">Full path where the themed workbook will be saved.</param>
        static void ApplySeasonalTheme(string inputPath, string outputPath)
        {
            // Load the existing workbook (uses the Workbook(string) constructor)
            Workbook workbook = new Workbook(inputPath);

            // Define a custom theme with 12 colors (example: Spring theme)
            Color[] seasonalColors = new Color[12]
            {
                Color.FromArgb(255, 255, 228, 225), // Background1 - MistyRose
                Color.FromArgb(255, 34, 139, 34),   // Text1 - ForestGreen
                Color.FromArgb(255, 255, 250, 240), // Background2 - FloralWhite
                Color.FromArgb(255, 85, 107, 47),   // Text2 - DarkOliveGreen
                Color.FromArgb(255, 60, 179, 113),  // Accent1 - MediumSeaGreen
                Color.FromArgb(255, 46, 139, 87),   // Accent2 - SeaGreen
                Color.FromArgb(255, 144, 238, 144), // Accent3 - LightGreen
                Color.FromArgb(255, 152, 251, 152), // Accent4 - PaleGreen
                Color.FromArgb(255, 0, 128, 0),     // Accent5 - Green
                Color.FromArgb(255, 34, 139, 34),   // Accent6 - ForestGreen (duplicate for illustration)
                Color.FromArgb(255, 0, 0, 255),     // Hyperlink - Blue
                Color.FromArgb(255, 255, 0, 0)      // Followed Hyperlink - Red
            };

            // Apply the custom theme (uses Workbook.CustomTheme method)
            workbook.CustomTheme("SpringSeasonalTheme", seasonalColors);

            // Save the themed workbook (uses Workbook.Save(string) method)
            workbook.Save(outputPath);
        }
    }
}
