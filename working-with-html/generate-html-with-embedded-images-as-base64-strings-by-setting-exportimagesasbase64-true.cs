// Title: C# – Export Excel to HTML with Embedded Base64 Images using Aspose.Cells
// Description: Demonstrates how to create a workbook, optionally add a picture, set HtmlSaveOptions.ExportImagesAsBase64 to true, and save the sheet as a single HTML file with images encoded as Base64 strings.
// Keywords: Aspose.Cells HTML export C# | embed images base64 Aspose.Cells | ExportImagesAsBase64 example | .NET Excel to HTML with embedded pictures | base64 image embedding in HTML
// Common Searches: Aspose.Cells export HTML base64 images C# | HtmlSaveOptions ExportImagesAsBase64 sample | save Excel as HTML with embedded pictures .NET | C# convert workbook to HTML with base64 images
// Developer Intent: Produce an HTML representation of an Excel workbook where all images are inlined as Base64 data URIs.
// Use Cases: Generate self‑contained web reports that don’t rely on external image files. | Create HTML email bodies with spreadsheet graphics that render correctly in mail clients. | Store a portable HTML snapshot of a workbook in a database or CMS.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to HTML with ExportImagesAsBase64 enabled, handling missing image files gracefully. | Explain how to retrieve the generated HTML as a string instead of writing it to disk when ExportImagesAsBase64 is true.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, optionally add a picture, set HtmlSaveOptions.ExportImagesAsBase64 to true, and save the sheet as a single HTML file with images encoded as Base64 strings.
    public class HtmlExportBase64Demo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "example.jpg";

            // Add an image if the file exists
            if (File.Exists(imagePath))
            {
                worksheet.Pictures.Add(0, 0, imagePath);
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping image insertion.");
            }

            // Create HTML save options with images embedded as Base64
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true
            };

            // Save the workbook as HTML
            string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
