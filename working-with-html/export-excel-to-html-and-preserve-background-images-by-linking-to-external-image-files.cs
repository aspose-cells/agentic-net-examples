// Title: Export Excel to HTML with linked background images using Aspose.Cells for .NET
// AI Prompts: Write C# that opens an .xlsx workbook, turns off Base64 image encoding in HtmlSaveOptions, creates a directory for picture files, and saves the workbook as an HTML page with the images referenced externally. | Demonstrate how to configure Aspose.Cells to export background graphics as individual image files during Excel‑to‑HTML conversion in .NET.
// Common Searches: Aspose.Cells C# export Excel to HTML without embedding images | Save Excel background pictures as separate files when converting to HTML using Aspose.Cells | How to disable Base64 image embedding in HtmlSaveOptions | Specify output folder for images in Aspose.Cells HTML conversion
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false | Aspose.Cells ImageFolder property for HTML export | C# Excel to HTML conversion with external image files | preserve workbook background pictures in HTML output | Aspose.Cells linked images in generated HTML

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an input.xlsx workbook, configures HtmlSaveOptions to export images as separate files, creates an images folder, sets the ImageFolder property when available, and saves the workbook as output.html with all background images stored externally.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputHtml = "output.html";
            const string imagesFolder = "images";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Configure HTML save options
            var options = new HtmlSaveOptions();
            options.ExportImagesAsBase64 = false; // export images as separate files

            // Ensure the images folder exists
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            // Set the folder for exported images if the property exists (newer API versions)
            var imageFolderProp = options.GetType().GetProperty("ImageFolder");
            if (imageFolderProp != null && imageFolderProp.CanWrite)
            {
                imageFolderProp.SetValue(options, imagesFolder);
            }

            // Save the workbook as HTML
            workbook.Save(outputHtml, options);
            Console.WriteLine($"Workbook saved to '{outputHtml}'. Images are stored in '{imagesFolder}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
