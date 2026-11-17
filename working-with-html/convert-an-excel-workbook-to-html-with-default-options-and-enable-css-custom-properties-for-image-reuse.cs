// Title: Convert an Excel workbook to HTML with external image files using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML with images saved as separate files | How to disable base64 image embedding when saving a workbook as HTML in .NET | Save Excel workbook as HTML and reuse images via CSS custom properties using Aspose.Cells | Generate HTML from .xlsx with external image resources using Aspose.Cells API | C# Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false example
// Tags: Aspose.Cells HtmlSaveOptions external image export | C# convert Excel workbook to HTML Aspose.Cells | HTML output with separate image files .NET | Disable base64 image embedding Aspose.Cells | CSS custom properties for HTML images Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRunner
{
    // This example loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to turn off Base64 image embedding, and saves the workbook as HTML so that images are written as separate files that can be referenced via CSS custom properties.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the Excel workbook from the file
                var workbook = new Workbook(inputPath);

                // Create HTML save options with default settings
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export images as separate files (not Base64) for CSS reuse
                    ExportImagesAsBase64 = false
                };

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
