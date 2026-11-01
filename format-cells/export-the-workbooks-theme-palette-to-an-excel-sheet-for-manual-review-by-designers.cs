// Title: How to export an Excel workbook’s theme palette to a new worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to list all theme colors from an Excel file using Aspose.Cells | Export Excel theme palette to a separate sheet for designer review in .NET | Retrieve ThemeColorType values and their ARGB and hex representations with Aspose.Cells | Add a ThemePalette worksheet and auto‑fit columns via Aspose.Cells API | Save a workbook that includes extracted theme colors to a new file in C#
// Tags: Aspose.Cells export theme colors to new sheet | GetThemeColor ThemeColorType enumeration C# | populate Excel sheet with ARGB and hex values | auto‑fit columns after data insertion Aspose.Cells | save workbook with added ThemePalette worksheet

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for ThemeColorType enum

// The example loads an existing Excel workbook, creates a new worksheet named "ThemePalette", writes headers, iterates through every ThemeColorType enum value, obtains each theme color with Workbook.GetThemeColor, records the color name, ARGB integer, and hex string, auto‑fits the columns, and saves the updated workbook to a new file for designers to review.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "WorkbookWithThemePalette.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet to hold the theme palette
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet paletteSheet = workbook.Worksheets[sheetIndex];
            paletteSheet.Name = "ThemePalette";

            // Write header row
            paletteSheet.Cells[0, 0].PutValue("Theme Color");
            paletteSheet.Cells[0, 1].PutValue("ARGB");
            paletteSheet.Cells[0, 2].PutValue("Hex");

            // Iterate through all ThemeColorType values and retrieve colors via Workbook API
            int row = 1;
            foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the color for the current theme element
                Color color = workbook.GetThemeColor(colorType);

                // Write the color information to the sheet
                paletteSheet.Cells[row, 0].PutValue(colorType.ToString());
                paletteSheet.Cells[row, 1].PutValue(color.ToArgb());
                paletteSheet.Cells[row, 2].PutValue(ColorTranslator.ToHtml(color));

                row++;
            }

            // Auto-fit columns for better readability
            paletteSheet.AutoFitColumns();

            // Save the workbook with the new sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
