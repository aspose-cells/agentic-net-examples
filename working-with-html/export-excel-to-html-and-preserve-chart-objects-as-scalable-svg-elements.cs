// Title: Export Excel workbook to HTML with charts saved as external image files using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML without embedding images | How to save Excel chart images as separate files when converting to HTML with Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 false example in .NET | Generate HTML from an .xlsx file and keep chart graphics as external PNG files using Aspose.Cells | C# convert Excel to HTML while preserving chart objects as external images
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false | C# export Excel to HTML external chart images | Aspose.Cells HTML conversion separate image files | Excel chart image extraction Aspose.Cells .NET | disable Base64 image embedding Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The program verifies the input .xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions to set ExportImagesAsBase64 = false so charts and other images are saved as separate image files, and then writes the workbook out as an HTML page.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Save chart and other images as separate files (not embedded as Base64)
                ExportImagesAsBase64 = false
            };

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
