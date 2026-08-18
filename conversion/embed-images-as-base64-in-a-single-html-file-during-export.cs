// Title: C# – Export Excel to a Single HTML File with Images Embedded as Base64 using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a picture, and save it as one self‑contained HTML document where all images are encoded as Base64 data URIs. Includes error handling for missing image files.
// Keywords: Aspose.Cells C# HTML export | embed images base64 Aspose.Cells | Save workbook as single HTML | HtmlSaveOptions ExportImagesAsBase64 | Excel to HTML conversion .NET | self‑contained HTML report
// Common Searches: Aspose.Cells embed picture as base64 in HTML C# | export Excel to single HTML file with inline images | HtmlSaveOptions SaveAsSingleFile example .NET | convert spreadsheet to HTML with embedded images | C# code to embed images in HTML using Aspose.Cells
// Developer Intent: Generate a standalone HTML version of an Excel workbook where every picture is inlined as a Base64 string, eliminating external image files.
// Use Cases: Send a complete HTML report via email without attaching separate image files. | Host an offline‑viewable spreadsheet on a static website where all graphics are bundled inside the page. | Create a portable product catalog from Excel where each photo is embedded directly in the HTML.
// AI Prompts: Show how to loop through a folder and add multiple pictures to a worksheet, then export all of them as Base64 in a single HTML file. | Provide code to compress images before embedding them as Base64 while still using SaveAsSingleFile. | Explain best practices for handling missing or corrupted image files when adding pictures to a worksheet prior to HTML export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEmbedImagesBase64
{
    // Demonstrates how to create a workbook, insert a picture, and save it as one self‑contained HTML document where all images are encoded as Base64 data URIs. Includes error handling for missing image files.
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

                // Path to the image to embed
                string imagePath = "example.jpg";

                // Add the image only if the file exists
                if (File.Exists(imagePath))
                {
                    // The image will be embedded as a Base64 string in the resulting HTML
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file \"{imagePath}\" not found. Skipping image insertion.");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Embed images directly as Base64 data URIs
                    ExportImagesAsBase64 = true,
                    // Save the entire workbook as a single HTML file (no external resources)
                    SaveAsSingleFile = true
                };

                // Save the workbook as HTML with embedded images
                string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file with embedded Base64 images saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
