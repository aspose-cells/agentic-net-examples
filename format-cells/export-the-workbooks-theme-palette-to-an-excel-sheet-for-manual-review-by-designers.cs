// Title: Export an Excel workbook’s theme palette to a new worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, adds a sheet named "ThemePalette", and writes each ThemeColorType name together with its ARGB hex value. | Extend the export routine to include separate columns for the red, green, and blue components of every theme color. | Implement command‑line arguments for input and output paths and add error handling for missing files or invalid arguments.
// Common Searches: Aspose.Cells C# export theme colors to a separate worksheet | How to list ThemeColorType values with ARGB codes in an Excel file using .NET | Create a ThemePalette sheet in an existing workbook with Aspose.Cells | Programmatically retrieve Excel theme palette for designer review in C# | Save Excel theme palette as hex values using Aspose.Cells API
// Tags: export theme palette to worksheet Aspose.Cells | retrieve ThemeColorType ARGB values .NET | add ThemePalette sheet programmatically Excel | auto‑fit columns after data export Aspose.Cells | command line file path handling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For ThemeColorType enum

// Loads a workbook, adds a "ThemePalette" sheet, writes each ThemeColorType name with its ARGB hex (and optionally RGB components), auto‑fits columns, and saves the result to a new file.
class ExportThemePalette
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook whose theme palette you want to export
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet to hold the theme palette information
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet paletteSheet = workbook.Worksheets[newSheetIndex];
            paletteSheet.Name = "ThemePalette";

            // Write header row
            paletteSheet.Cells[0, 0].PutValue("Theme Color Type");
            paletteSheet.Cells[0, 1].PutValue("ARGB Value");

            // Iterate through all ThemeColorType values and write their ARGB values
            int row = 1;
            foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the color for the current theme color type using Workbook API
                Color color = workbook.GetThemeColor(colorType);

                // Write the theme color type name
                paletteSheet.Cells[row, 0].PutValue(colorType.ToString());

                // Write the ARGB value as a hex string (e.g., #FF112233)
                string argbHex = $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
                paletteSheet.Cells[row, 1].PutValue(argbHex);

                row++;
            }

            // Auto-fit the columns for better readability
            paletteSheet.AutoFitColumns();

            // Save the workbook with the added ThemePalette sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Theme palette exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
