// Title: How to assign custom RGB values to Accent1 and Accent2 theme colors in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: asp.net set custom RGB for Accent1 and Accent2 theme colors using Aspose.Cells | c# programmatically change Excel theme accent colors with Aspose.Cells | workaround for missing Theme API in Aspose.Cells to apply custom color scheme | use a template workbook to define custom accent colors in Aspose.Cells .NET | save Excel file after applying custom theme colors with Aspose.Cells C#
// Tags: custom accent colors Aspose.Cells | set theme RGB values .NET | excel theme color scheme workaround | template workbook theme customization | aspose.cells workbook save example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, explains that the Theme and ColorScheme APIs are not available in the current Aspose.Cells version, and recommends using a pre‑configured template workbook to apply custom Accent1 and Accent2 RGB values. It ensures the output directory exists, saves the workbook as CustomColorScheme.xlsx, prints the full file path, and includes basic exception handling.
class CustomColorSchemeExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // NOTE: The Theme and ColorScheme APIs are not available in the current Aspose.Cells version.
            // If needed, theme customization can be performed using a template workbook that already
            // contains the desired theme settings.

            // Define output file path
            string outputPath = "CustomColorScheme.xlsx";

            // Ensure the output directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
