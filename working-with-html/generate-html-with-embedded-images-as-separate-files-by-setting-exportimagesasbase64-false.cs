// Title: Export Excel to HTML with External Image Files using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert a picture if the file exists, set HtmlSaveOptions.ExportImagesAsBase64 to false, and save the workbook as HTML so that each picture is written to a separate image file instead of being embedded as Base64.
// Keywords: Aspose.Cells | HTML export | ExportImagesAsBase64 false | external image files | C# .NET | Excel to HTML | picture insertion | separate image files | web‑ready report | image caching
// Common Searches: Aspose.Cells export Excel to HTML external images | C# HtmlSaveOptions ExportImagesAsBase64 false example | Save workbook as HTML with separate image files | Prevent Base64 images in Aspose.Cells HTML export | Insert picture into worksheet and export to HTML Aspose.Cells
// Developer Intent: Generate HTML from a workbook where embedded pictures are saved as individual image files rather than Base64 strings.
// Use Cases: Build web reports that reference logos and charts as separate files for caching and easy updates. | Automate conversion of product‑catalog Excel sheets to HTML pages that link to external images for faster load times. | Create HTML email templates from Excel data while keeping images as linked files to meet email client restrictions. | Develop a documentation portal that converts Excel manuals to HTML with images stored in a dedicated assets folder.
// AI Prompts: Write C# code with Aspose.Cells to convert an Excel workbook to HTML, exporting pictures as separate files and handling missing image paths gracefully. | Explain which HtmlSaveOptions properties must be configured to stop Base64 image embedding when saving a workbook as HTML. | Suggest modifications to the sample that allow specifying a custom output folder for the exported image files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, insert a picture if the file exists, set HtmlSaveOptions.ExportImagesAsBase64 to false, and save the workbook as HTML so that each picture is written to a separate image file instead of being embedded as Base64.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image to be inserted
                string imagePath = "example.jpg";

                // Add the image only if the file exists
                if (File.Exists(imagePath))
                {
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // Configure HTML save options to export images as separate files
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = false // Do not embed images as Base64
                };

                // Save the workbook as HTML; images will be saved as separate files
                string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML saved with images as separate files: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
