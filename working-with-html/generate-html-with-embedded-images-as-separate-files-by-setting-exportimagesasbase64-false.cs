// Title: Export Excel workbook to HTML with external images using Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a PNG picture, configure HtmlSaveOptions.ExportImagesAsBase64 = false, and save the workbook so each picture is written as a separate file next to the generated HTML.
// Keywords: Aspose.Cells | .NET | HTML export | external images | ExportImagesAsBase64 false | C# workbook to HTML | picture insertion | separate image files | global | US
// Common Searches: Aspose.Cells export HTML external images C# | HtmlSaveOptions ExportImagesAsBase64 example | Save Excel as HTML with image files | C# generate HTML from workbook without base64 | keep pictures separate when converting Excel to HTML
// Developer Intent: Generate an HTML file from an Excel workbook while storing inserted pictures as independent image files rather than embedding them as Base64 strings.
// Use Cases: Publish web‑ready reports where charts and graphics are cached as separate assets for faster loading. | Create HTML email templates from Excel data, keeping images external to reduce email size and simplify asset management. | Automate bulk conversion of spreadsheets to static web pages, ensuring each picture is saved as an individual file for easy deployment.
// AI Prompts: Write C# code with Aspose.Cells that exports a workbook to HTML, disables Base64 image embedding, and saves pictures to a specified folder. | Provide a step‑by‑step tutorial for inserting a PNG into a worksheet and exporting the workbook to HTML with external image files using Aspose.Cells for .NET. | Explain how to detect a missing picture file and skip insertion gracefully when converting an Excel workbook to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to create a workbook, insert a PNG picture, configure HtmlSaveOptions.ExportImagesAsBase64 = false, and save the workbook so each picture is written as a separate file next to the generated HTML.
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
                string imagePath = "sample_image.png";

                // Add the image only if the file exists
                if (File.Exists(imagePath))
                {
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Skipping picture insertion.");
                }

                // Configure HTML save options to export images as separate files
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = false // Do not embed images as Base64
                };

                // Save the workbook as HTML; images will be saved as separate files alongside the HTML
                string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML exported successfully to \"{outputPath}\" with images saved as separate files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
