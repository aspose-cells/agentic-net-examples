// Title: C# – Export Aspose.Cells Workbook to a Single HTML File with Embedded Base64 Images
// Description: Shows how to create a workbook, add a picture, and use HtmlSaveOptions (ExportImagesAsBase64 = true, SaveAsSingleFile = true) to produce a self‑contained HTML file, including handling of missing image files.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportImagesAsBase64 | SaveAsSingleFile | embed images base64 | single HTML export | Excel to HTML | self-contained HTML | picture insertion | workbook to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells export HTML with base64 images C# | Save Excel workbook as single HTML file Aspose | Embed pictures as Base64 in Aspose.Cells HTML output | C# generate self‑contained HTML from Excel using Aspose | HtmlSaveOptions ExportImagesAsBase64 example
// Developer Intent: Generate a self‑contained HTML document from an Excel workbook with all images encoded as Base64.
// Use Cases: Create an email‑ready HTML report that includes logos or charts without external files. | Provide a portable web preview of a spreadsheet where every resource is bundled in one file. | Distribute a static HTML version of a workbook for environments without file system access.
// AI Prompts: Modify the example to apply a custom CSS stylesheet while still embedding images as Base64. | Show how to export multiple worksheets into one HTML file with Base64 images using Aspose.Cells. | Explain strategies for reducing HTML size when large images are embedded as Base64 strings.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a picture, and use HtmlSaveOptions (ExportImagesAsBase64 = true, SaveAsSingleFile = true) to produce a self‑contained HTML file, including handling of missing image files.
    public class HtmlExportBase64Demo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image to be inserted
                string imagePath = "example.jpg";

                // Add the image only if the file exists
                if (File.Exists(imagePath))
                {
                    worksheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Set up HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Embed images directly as Base64 strings in the HTML
                    ExportImagesAsBase64 = true,
                    // Save the entire workbook as a single HTML file (no external resources)
                    SaveAsSingleFile = true
                };

                // Export the workbook to HTML with the configured options
                string outputPath = "output.html";
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
