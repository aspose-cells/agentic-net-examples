// Title: Export an Aspose.Cells workbook to a single self‑contained HTML file with embedded Base64 images in C#
// AI Prompts: Write C# code that creates a workbook, inserts a picture, and saves it as one HTML file with all images encoded as Base64 using Aspose.Cells. | Demonstrate how to set HtmlSaveOptions.ExportImagesAsBase64 and SaveAsSingleFile for exporting a workbook to HTML in C#. | Add a file‑existence check before inserting a picture into a worksheet and then export the workbook to a single HTML page.
// Common Searches: Aspose.Cells C# export workbook to single HTML file with base64 images | how to embed pictures as Base64 in HTML using Aspose.Cells HtmlSaveOptions | save Excel as self‑contained HTML in .NET without external image files | C# example for ExportImagesAsBase64 and SaveAsSingleFile in Aspose.Cells | convert Excel to one HTML page with embedded images using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | single file HTML export Aspose.Cells | embed images base64 Aspose.Cells C# | Excel to self-contained HTML Aspose.Cells | picture insertion worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, optionally insert an image, and export the workbook as a single HTML file with all pictures embedded as Base64 strings by configuring HtmlSaveOptions.
    public class EmbedImagesBase64SingleHtml
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some text to identify the image location
                sheet.Cells["A1"].PutValue("Image embedded as Base64");

                // Path to the image to embed
                string imagePath = "example.jpg";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Insert the image into the worksheet at row 2, column 0 (C2)
                    sheet.Pictures.Add(2, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{imagePath}' not found. Skipping image insertion.");
                }

                // Set up HTML save options
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    // Export images directly as Base64 strings inside the <img> tags
                    ExportImagesAsBase64 = true,
                    // Save the entire workbook as a single HTML file (no external resources)
                    SaveAsSingleFile = true
                };

                // Export the workbook to HTML with the configured options
                string outputPath = "output.html";
                workbook.Save(outputPath, options);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EmbedImagesBase64SingleHtml.Run();
        }
    }
}
