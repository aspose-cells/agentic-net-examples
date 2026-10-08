// Title: Convert an Excel workbook containing WordArt to W3C‑validated HTML with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file with WordArt, configures HtmlSaveOptions to embed all images as Base64 and export only the active worksheet, then saves the result as HTML that passes W3C CSS gradient validation. | Show how to add file‑existence checking and robust exception handling around an Aspose.Cells workbook‑to‑HTML conversion that includes WordArt content. | Demonstrate setting Aspose.Cells HtmlSaveOptions so that the generated CSS for gradient fills complies with W3C standards.
// Common Searches: aspnet convert excel with wordart to html base64 images w3c css validation | c# Aspose.Cells export active worksheet only as html gradient fill | how to embed wordart images as base64 when saving excel to html using Aspose.Cells | validate css gradients in html output from Aspose.Cells workbook conversion
// Tags: Aspose.Cells HtmlSaveOptions Base64 image embedding | export active worksheet to html Aspose.Cells | WordArt gradient CSS W3C validation | excel to html conversion C# Aspose.Cells | Base64 image embedding in html Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // This program checks for the presence of an Excel file that contains WordArt, loads it with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64 and export only the active sheet, then saves the workbook as an HTML file that adheres to W3C CSS gradient standards.
class Program
{
    static void Main()
    {
        try
        {
            const string sourceFile = "WordArtWorkbook.xlsx";
            const string outputFile = "WordArtWorkbook.html";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Error: The file '{sourceFile}' was not found.");
                return;
            }

            // Load the workbook that contains WordArt
            Workbook workbook = new Workbook(sourceFile);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Embed images (including gradient images) as Base64 strings
                ExportImagesAsBase64 = true,
                // Export only the active worksheet (optional)
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
