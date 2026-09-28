// Title: Export Excel to HTML with external image files using Aspose.Cells for .NET (ExportImagesAsBase64 = false)
// AI Prompts: Write C# code that loads an .xlsx workbook and saves it as an HTML file with each image written to a separate file by setting HtmlSaveOptions.ExportImagesAsBase64 to false. | Show how to define a custom folder for the exported images when converting an Excel workbook to HTML with Aspose.Cells.
// Common Searches: how to save Excel workbook as HTML with images in separate files using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false example | C# export Excel to HTML external image files Aspose.Cells tutorial
// Tags: Aspose.Cells HtmlSaveOptions disable base64 image export | export excel to html external images C# | set ExportImagesFolder Aspose.Cells | html conversion separate image files Aspose.Cells | disable base64 image embedding Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The program loads an Excel workbook, configures HtmlSaveOptions to turn off Base64 image embedding (ExportImagesAsBase64 = false) and optionally sets an ExportImagesFolder, then saves the workbook as HTML where each image is saved as an external file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export images as separate files instead of embedding them as Base64 strings
                ExportImagesAsBase64 = false,

                // Optional: specify the folder where images will be saved
                // By default, images are saved in a subfolder named after the HTML file
                // ExportImagesFolder = "Images"
            };

            // Save the workbook as HTML. Images will be written to separate files.
            workbook.Save("Output.html", saveOptions);
        }
    }
}
