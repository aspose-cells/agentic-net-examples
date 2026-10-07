// Title: Export an Excel workbook to HTML with external image files and preserve worksheet background images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file, configures HtmlSaveOptions to save the workbook as HTML, stores all embedded images and worksheet background images in a separate folder, and disables base‑64 image embedding. | Show how to use reflection in C# to set HtmlSaveOptions.ExportImagesFolder and HtmlSaveOptions.ExportBackgroundImages only when those properties exist in the installed Aspose.Cells version.
// Common Searches: Aspose.Cells .NET export Excel to HTML with images saved as separate files | how to keep worksheet background picture when converting Excel to HTML using Aspose.Cells | HtmlSaveOptions ExportImagesFolder property example | C# save workbook as HTML without base64 images Aspose.Cells | export background images from Excel to HTML using Aspose.Cells reflection
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesFolder | Aspose.Cells export background images to HTML | C# save Excel as HTML external images | reflection for HtmlSaveOptions compatibility | preserve worksheet background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, configures HtmlSaveOptions to write images and worksheet background images to a specified folder as separate files (not base64), and saves the workbook as an HTML document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";
            const string imagesFolder = "Images";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = false,          // Export images as separate files
                ExportActiveWorksheetOnly = false      // Export all worksheets
            };

            // Set the folder for external images if the property exists in the current version
            var exportImagesFolderProp = typeof(HtmlSaveOptions).GetProperty("ExportImagesFolder");
            if (exportImagesFolderProp != null && exportImagesFolderProp.CanWrite)
            {
                exportImagesFolderProp.SetValue(htmlOptions, imagesFolder);
            }

            // Enable background image export if the property exists
            var exportBgImagesProp = typeof(HtmlSaveOptions).GetProperty("ExportBackgroundImages");
            if (exportBgImagesProp != null && exportBgImagesProp.CanWrite)
            {
                exportBgImagesProp.SetValue(htmlOptions, true);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
