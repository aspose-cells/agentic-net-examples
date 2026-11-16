// Title: Convert HTML to XLSX in C# while preserving rows and columns hidden by CSS display:none using Aspose.Cells
// Common Searches: C# Aspose.Cells convert HTML file to Excel keep display:none rows hidden | How to retain hidden columns from HTML when saving as XLSX with Aspose.Cells | Preserve CSS display:none elements during HTML to XLSX conversion in .NET | Aspose.Cells load HTML and export to Excel without showing hidden rows | Convert HTML tables to Excel in C# while maintaining hidden visibility
// Tags: Aspose.Cells HTML import with CSS visibility handling | C# convert HTML to Excel while retaining hidden elements | HTML to XLSX export preserving display:none styling | Aspose.Cells workbook save XLSX with hidden row support | Excel generation from HTML respecting CSS hidden columns

using Aspose.Cells;
using System;
using System.IO;

// The example loads an HTML file into an Aspose.Cells Workbook in C#, ensures the output folder exists, and saves the workbook as an XLSX file while preserving rows and columns that are hidden via CSS display:none.
class Program
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a workbook (default options)
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file (XLSX format)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
