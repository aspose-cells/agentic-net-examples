// Title: Embed Worksheet Images as Base64 in HTML with Aspose.Cells (C#)
// Description: Creates a workbook, optionally inserts a picture into cell A1, sets HtmlSaveOptions.ExportImagesAsBase64 to true, and saves the file as a single self‑contained HTML document where the image is stored as a Base64 string.
// Keywords: Aspose.Cells ExportImagesAsBase64 | C# embed images in HTML | self‑contained HTML workbook | Base64 image embedding .NET | Aspose.Cells HTML export options
// Common Searches: Aspose.Cells ExportImagesAsBase64 example C# | how to embed pictures in HTML output using Aspose.Cells | save Excel as single HTML file with embedded images | C# convert workbook to HTML with Base64 images
// Developer Intent: Produce an HTML file from an Excel workbook where all worksheet images are encoded as Base64 and included directly in the markup.
// Use Cases: Generate a portable HTML report that displays charts and pictures without external files. | Create email‑ready HTML content from Excel where embedded images guarantee proper rendering. | Distribute offline documentation that contains visual elements without requiring image folders.
// AI Prompts: Show how to set a custom HTML title while using ExportImagesAsBase64 in Aspose.Cells. | Provide a C# snippet that loads an existing .xlsx, embeds its images as Base64, and writes the HTML to a MemoryStream. | Explain techniques to reduce memory usage when exporting large images as Base64 with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, optionally inserts a picture into cell A1, sets HtmlSaveOptions.ExportImagesAsBase64 to true, and saves the file as a single self‑contained HTML document where the image is stored as a Base64 string.
    public class ExportImagesAsBase64Demo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string imagePath = "example.jpg";

                // Verify the image file exists before adding it to the worksheet
                if (File.Exists(imagePath))
                {
                    // Add the image at cell A1 (row 0, column 0)
                    worksheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Set HTML save options to embed images as Base64 strings
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                string outputPath = "output.html";
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved as HTML with embedded images to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportImagesAsBase64Demo.Run();
        }
    }
}
