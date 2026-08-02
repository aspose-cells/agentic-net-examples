// Title: Batch Apply a Custom Holiday Theme to Multiple Excel Workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console app that scans an input folder, loads each .xlsx/.xls/.xlsm file with Aspose.Cells, creates a "HolidayTheme" using twelve predefined colors, applies the theme to every workbook, and saves the themed copies to an output folder while logging errors.
// Keywords: Aspose.Cells batch theme | C# apply custom Excel theme | holiday color scheme Excel | multiple workbook processing .NET | custom theme programmatically | Excel calendar theming | automate Excel styling | Aspose.Cells CustomTheme method
// Common Searches: How to batch apply a custom theme to Excel files using Aspose.Cells C# | Apply holiday colors to many workbooks programmatically | Aspose.Cells example for processing multiple Excel workbooks | C# script to add a seasonal theme to all spreadsheets in a folder | Automate Excel theme changes with Aspose.Cells .NET
// Developer Intent: Automatically add a predefined holiday color theme to every Excel workbook in a directory and save the themed versions.
// Use Cases: Create holiday‑themed financial reports by applying a consistent color palette to all monthly workbooks before distribution. | Prepare seasonal marketing dashboards with uniform branding across dozens of spreadsheets using a single batch run. | Update employee schedule calendars so each workbook reflects the holiday color scheme without manual editing.
// AI Prompts: Generate C# code that uses Aspose.Cells to define a custom theme named "HolidayTheme" with specific RGB colors and apply it to a workbook. | Write a method that iterates through a folder, loads each .xlsx/.xls/.xlsm file, applies the HolidayTheme via CustomTheme, saves the file to an output directory, and logs any failures. | Explain how to extend the batch processor to also modify existing named styles in each workbook to match the holiday theme colors.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// A C# console app that scans an input folder, loads each .xlsx/.xls/.xlsm file with Aspose.Cells, creates a "HolidayTheme" using twelve predefined colors, applies the theme to every workbook, and saves the themed copies to an output folder while logging errors.
class HolidayThemeBatchProcessor
{
    // Define the holiday theme name
    private const string HolidayThemeName = "HolidayTheme";

    // Define 12 colors for the theme (Background1, Text1, Background2, Text2, Accent1-6, Hyperlink, FollowedHyperlink)
    private static readonly Color[] HolidayColors = new Color[]
    {
        Color.FromArgb(255, 255, 255), // Background1 - White
        Color.FromArgb(0, 0, 0),       // Text1 - Black
        Color.FromArgb(255, 228, 196), // Background2 - Bisque (warm)
        Color.FromArgb(0, 100, 0),     // Text2 - DarkGreen
        Color.FromArgb(220, 20, 60),   // Accent1 - Crimson (red)
        Color.FromArgb(34, 139, 34),   // Accent2 - ForestGreen
        Color.FromArgb(255, 215, 0),   // Accent3 - Gold
        Color.FromArgb(255, 140, 0),   // Accent4 - DarkOrange
        Color.FromArgb(138, 43, 226),  // Accent5 - BlueViolet
        Color.FromArgb(70, 130, 180),  // Accent6 - SteelBlue
        Color.FromArgb(0, 0, 255),     // Hyperlink - Blue
        Color.FromArgb(128, 0, 128)    // FollowedHyperlink - Purple
    };

    static void Main()
    {
        // Input folder containing workbooks to process
        string inputFolder = @"C:\Workbooks\Input";
        // Output folder where themed workbooks will be saved
        string outputFolder = @"C:\Workbooks\Output";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each Excel file in the input folder (supports .xlsx, .xls, .xlsm)
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                continue; // Skip non-Excel files

            // Verify the file actually exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply the custom holiday theme
                workbook.CustomTheme(HolidayThemeName, HolidayColors);

                // Build output file path (preserve original name)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the themed workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Applied holiday theme to: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
