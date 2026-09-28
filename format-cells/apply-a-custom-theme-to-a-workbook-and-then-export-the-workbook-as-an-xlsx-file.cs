// Title: How to apply a custom .thmx theme to an Aspose.Cells workbook and save it as an XLSX file in C#
// AI Prompts: Write C# code that loads a .thmx theme file, uses Workbook.CustomTheme to apply it to a new Aspose.Cells workbook, and then saves the workbook in XLSX format. | Demonstrate error handling for missing theme files when applying a custom theme with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells apply .thmx theme to new workbook | Saving a themed workbook as XLSX with Aspose.Cells .NET | How to validate theme file existence before using Workbook.CustomTheme | Example of Workbook.CustomTheme method usage in Aspose.Cells | Exporting a workbook with custom theme to XLSX using C#
// Tags: custom thmx theme application Aspose.Cells | Workbook.CustomTheme method C# | save themed workbook as xlsx Aspose.Cells | theme file validation .NET | export workbook with custom theme

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The program checks for a customTheme.thmx file, creates a new Aspose.Cells Workbook, applies the theme via Workbook.CustomTheme, and saves the result as CustomThemeWorkbook.xlsx in XLSX format, with basic exception handling for missing files and other errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the custom theme file.
            string themePath = "customTheme.thmx";

            // Verify that the theme file exists to avoid FileNotFoundException.
            if (!File.Exists(themePath))
            {
                Console.WriteLine($"Theme file not found: {themePath}");
                return;
            }

            // Create a new workbook (empty workbook with a default worksheet).
            Workbook workbook = new Workbook();

            // Apply the custom theme to the workbook. Passing null for colors uses the theme's default colors.
            workbook.CustomTheme(themePath, null);

            // Export the workbook as an XLSX file.
            string outputPath = "CustomThemeWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
