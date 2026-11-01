// Title: Convert an Excel workbook to HTML with external image files using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportImagesAsBase64 to false, and saves the workbook as an HTML file with images written to a separate folder. | Show how to create the images directory, configure Aspose.Cells HtmlSaveOptions for external image export, and handle missing input workbook errors in C#. | Provide a step‑by‑step example of converting Excel to HTML while disabling Base64 image embedding and ensuring the images are saved alongside the HTML output.
// Common Searches: Aspose.Cells C# export Excel to HTML without embedding images as Base64 | How to save images as separate files when converting .xlsx to HTML using Aspose.Cells | C# HtmlSaveOptions ExportImagesAsBase64 false example for Aspose.Cells | Generate HTML and image folder from Excel workbook with Aspose.Cells .NET | Disable Base64 image embedding in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions disable Base64 images | C# export Excel workbook to HTML with image folder | generate separate image files during HTML export Aspose.Cells | Excel to HTML conversion without embedded images .NET | configure Aspose.Cells image export path

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel file, ensures an images directory exists, configures HtmlSaveOptions with ExportImagesAsBase64 = false, and saves the workbook as HTML, causing Aspose.Cells to write each picture to a separate image file in the folder.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output paths
            string inputPath = "input.xlsx";
            string outputHtml = "output.html";
            string imagesFolder = "Images";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the images folder exists (used by Aspose.Cells when exporting images)
            if (!Directory.Exists(imagesFolder))
            {
                Directory.CreateDirectory(imagesFolder);
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Save images as separate files (not Base64)
                ExportImagesAsBase64 = false
                // Note: ExportImageFolderPath is not available in this version;
                // Aspose.Cells will place images in a folder named after the HTML file.
            };

            // Save the workbook as HTML using the configured options
            workbook.Save(outputHtml, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputHtml}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
