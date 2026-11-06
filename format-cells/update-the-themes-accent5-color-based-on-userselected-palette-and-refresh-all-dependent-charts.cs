// Title: Set the Accent5 color of an Excel workbook’s theme and refresh dependent charts using Aspose.Cells for .NET
// Common Searches: c# asp.net change excel theme accent5 color with aspose.cells | how to update excel theme palette and refresh linked charts programmatically | asp.net modify workbook theme color and recalculate charts using aspose.cells | apply custom accent5 shade to excel file and refresh charts in .net | update excel theme colors dynamically and refresh dependent visuals with aspose.cells
// Tags: set workbook theme accent5 color Aspose.Cells | update excel theme palette C# | refresh charts after theme modification Aspose.Cells | apply custom theme color to Excel workbook .NET | save workbook with modified theme Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, changes its theme Accent5 color to a user‑specified System.Drawing.Color via Workbook.SetThemeColor, ensures the output directory exists, saves the modified file, and logs any errors. Updating the theme automatically refreshes charts that depend on the theme colors.
class ThemeUpdater
{
    // Updates the workbook's theme Accent5 color.
    public static void UpdateAccent5(string inputFilePath, string outputFilePath, Color newAccent5Color)
    {
        try
        {
            // Verify the source file exists.
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputFilePath);

            // Set the Accent5 color in the theme.
            workbook.SetThemeColor(ThemeColorType.Accent5, newAccent5Color);

            // Ensure the destination directory exists.
            string destDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Save the modified workbook.
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error updating theme: {ex.Message}");
        }
    }

    // Example usage.
    static void Main()
    {
        // Path to the source workbook.
        string sourcePath = @"C:\Docs\Sample.xlsx";

        // Path where the updated workbook will be saved.
        string destinationPath = @"C:\Docs\Sample_Updated.xlsx";

        // User‑selected palette color for Accent5 (e.g., a shade of teal).
        Color userSelectedAccent5 = Color.FromArgb(0, 128, 128);

        // Perform the update.
        UpdateAccent5(sourcePath, destinationPath, userSelectedAccent5);
    }
}
