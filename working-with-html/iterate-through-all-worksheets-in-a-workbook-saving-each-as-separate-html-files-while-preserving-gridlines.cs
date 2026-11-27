// Title: Save each worksheet of an Excel workbook as a separate HTML file with gridlines using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells loop through worksheets and export each to HTML with gridlines | How to generate individual HTML files for every sheet in an Excel file using Aspose.Cells | Saving Excel worksheets as separate HTML pages while preserving gridlines in .NET | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly example for multiple sheets | Create safe filenames for worksheet HTML export in C# Aspose.Cells
// Tags: export worksheet to HTML Aspose.Cells | HtmlSaveOptions ExportGridLines | batch convert Excel sheets to HTML C# | set active worksheet index Aspose.Cells | sanitize worksheet name for file system C# | save each sheet as separate HTML file Aspose | create output directory for HTML exports .NET

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an Excel workbook, creates an output folder, iterates through all worksheets, sets each as the active sheet, and saves it as an individual HTML file with gridlines enabled via HtmlSaveOptions, using a helper method to generate safe file names.
class Program
{
    static void Main(string[] args)
    {
        // Determine the input workbook path.
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        // Load the workbook.
        Workbook workbook;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Create an output directory for the HTML files.
        string outputDir = "HtmlSheets";
        try
        {
            Directory.CreateDirectory(outputDir);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating output directory '{outputDir}': {ex.Message}");
            return;
        }

        // Iterate through each worksheet and save it as an individual HTML file.
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            string sheetName = sheet.Name;
            string safeSheetName = GetSafeFileName(sheetName);
            string outputPath = Path.Combine(outputDir, $"{safeSheetName}.html");

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExportGridLines = true
            };

            // Set the current worksheet as the active one before saving.
            workbook.Worksheets.ActiveSheetIndex = i;

            try
            {
                // Save using the overload that accepts SaveOptions.
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Worksheet '{sheetName}' saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving worksheet '{sheetName}': {ex.Message}");
            }
        }
    }

    // Helper method to replace invalid filename characters.
    private static string GetSafeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
