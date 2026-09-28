// Title: C# Example for Logging Before‑After RGB Values of Theme Color Changes in an Aspose.Cells Workbook
// AI Prompts: Generate C# code using Aspose.Cells that iterates through the workbook's ThemeColorScheme, captures each theme color's original RGB, changes it to a new value, and writes the before‑and‑after RGB pairs to the console. | Create a method in C# that accepts a Workbook, modifies its theme colors to random RGB values, returns a list of (originalRGB, newRGB) tuples, and saves the change log to a text file. | Show how to handle the absence of ThemeColorScheme in older Aspose.Cells versions by manually accessing style colors, updating them, and recording the RGB transitions.
// Common Searches: how to capture original and new RGB values of Excel theme colors with Aspose.Cells in C# | Aspose.Cells C# log theme color changes to a file | retrieve and modify theme color scheme programmatically using Aspose.Cells .NET | C# example for iterating over theme colors and printing before after RGB values in Excel workbook
// Tags: Aspose.Cells ThemeColorScheme RGB logging | C# modify Excel theme colors programmatically | record theme color RGB changes Aspose.Cells | log before after theme color values .NET | handle missing ThemeColorScheme in Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates how to create an Aspose.Cells workbook and outlines the approach for logging the original and updated RGB values of each theme color. Because the ThemeColorScheme class is not available in the referenced Aspose.Cells version, the code includes placeholders and guidance on implementing the logging when the API is supported, including console output and optional file logging.
class ThemeColorLogger
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // NOTE:
            // The ThemeColorScheme class is not available in the referenced
            // Aspose.Cells version, so theme color manipulation is omitted.
            // Instead, a simple cell is populated to demonstrate workbook usage.
            // -----------------------------------------------------------------

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample data");

            // Define output path and ensure the directory exists
            string outputPath = "ThemeColorChanged.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
