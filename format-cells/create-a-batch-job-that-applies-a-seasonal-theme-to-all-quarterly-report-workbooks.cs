// Title: C# Batch Apply Spring Theme to Multiple Excel Workbooks Using Aspose.Cells
// Description: C# console app scans a folder of .xlsx quarterly reports, creates a 12‑color spring palette, applies it via Workbook.CustomTheme, and saves themed workbooks to an output folder with error handling.
// Keywords: Aspose.Cells | C# batch theme | custom Excel theme | seasonal color palette | Workbook.CustomTheme | spring theme Excel | bulk workbook processing | quarterly report automation | .NET Excel library | Excel theming script
// Common Searches: apply custom theme to multiple Excel files C# | Aspose.Cells batch theme example | Workbook.CustomTheme bulk processing | C# script to add seasonal palette to workbooks | automate Excel theme for quarterly reports
// Developer Intent: Automatically apply a spring‑style custom theme to every quarterly‑report workbook in a directory and write the themed copies to a separate folder.
// Use Cases: Generate spring‑themed versions of all quarterly reports for marketing or internal distribution without manual editing. | Update branding across an existing library of Excel reports when corporate colors change. | Create separate seasonal workbook sets (spring, summer, fall, winter) as part of an automated publishing pipeline.
// AI Prompts: Write C# code that uses Aspose.Cells to load each .xlsx file in a folder, apply a 12‑color spring palette as a custom theme, and save the result to an output directory. | Show how to call Workbook.CustomTheme with a named theme and a Color[] array while handling missing folders and file‑not‑found errors. | Explain best practices for batch processing Excel workbooks in .NET, including logging, exception handling, and performance tips.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace SeasonalThemeBatchJob
{
    // C# console app scans a folder of .xlsx quarterly reports, creates a 12‑color spring palette, applies it via Workbook.CustomTheme, and saves themed workbooks to an output folder with error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder that contains the quarterly report workbooks
            string inputFolder = @"C:\QuarterlyReports\Input";
            // Define the folder where the themed workbooks will be saved
            string outputFolder = @"C:\QuarterlyReports\Output";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Define a seasonal (e.g., Spring) theme color palette (12 colors)
            // Index mapping: 0-Background1, 1-Text1, 2-Background2, 3-Text2,
            // 4-Accent1, 5-Accent2, 6-Accent3, 7-Accent4, 8-Accent5, 9-Accent6,
            // 10-Hyperlink, 11-FollowedHyperlink
            Color[] springColors = new Color[]
            {
                Color.FromArgb(255, 228, 225), // Light pink background
                Color.FromArgb(34, 139, 34),   // Forest green text
                Color.FromArgb(255, 250, 240), // Ivory background
                Color.FromArgb(85, 107, 47),   // Dark olive text
                Color.FromArgb(60, 179, 113),  // Medium sea green accent
                Color.FromArgb(255, 182, 193), // Light pink accent
                Color.FromArgb(144, 238, 144), // Light green accent
                Color.FromArgb(255, 215, 0),   // Gold accent
                Color.FromArgb(173, 216, 230), // Light blue accent
                Color.FromArgb(221, 160, 221), // Plum accent
                Color.FromArgb(0, 0, 255),     // Blue hyperlink
                Color.FromArgb(128, 0, 128)    // Purple followed hyperlink
            };

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Apply the custom seasonal theme
                    workbook.CustomTheme("SpringTheme", springColors);

                    // Determine the output file path (same name, different folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the themed workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Applied Spring theme to: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
