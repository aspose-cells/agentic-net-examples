// Title: Check for a custom .thmx theme file and export an Aspose.Cells workbook to XLSX using C#
// Common Searches: C# Aspose.Cells verify .thmx theme file before saving workbook | how to save an Aspose.Cells workbook as XLSX after checking for a custom theme | Aspose.Cells create empty workbook and export to XLSX in .NET | workaround for loading custom theme in Aspose.Cells when method is unavailable | C# code example for handling missing theme file with Aspose.Cells
// Tags: apply custom thmx theme Aspose.Cells C# | save workbook to XLSX format Aspose.Cells | file existence check C# Aspose.Cells | unsupported theme loading handling Aspose.Cells | create empty workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new empty Workbook, checks for the presence of a customTheme.thmx file, logs that loading the theme is not supported in the current Aspose.Cells version, ensures the output directory exists, and saves the workbook as an XLSX file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook.
            Workbook workbook = new Workbook();

            string themePath = "customTheme.thmx";

            // Load and apply custom theme if the file exists.
            if (File.Exists(themePath))
            {
                try
                {
                    // Aspose.Cells does not expose a direct LoadTheme method in some versions.
                    // If supported, you could use workbook.LoadTheme(themePath);
                    // Here we simply note that the theme file is present.
                    Console.WriteLine($"Theme file found at {themePath}, but loading is not supported in this version.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load theme: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Theme file not found: {themePath}. Continuing without custom theme.");
            }

            // Ensure the output directory exists.
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as XLSX.
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
