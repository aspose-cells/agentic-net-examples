// Title: Generate HTML from an Aspose.Cells workbook with embedded Base64 images in C#
// AI Prompts: Create an HTML file from a Workbook and embed all worksheet pictures as Base64 strings using Aspose.Cells HtmlSaveOptions in C#. | Set HtmlSaveOptions.ExportImagesAsBase64 to true and save the workbook as HTML with inline image data URIs. | Add a picture to a worksheet, then export the workbook to HTML ensuring the image appears as a Base64‑encoded data URI.
// Common Searches: asp.net c# export excel to html with inline base64 images using aspose.cells | how to embed worksheet pictures as data URIs when saving as html with aspose cells | htmlsaveoptions exportimagesasbase64 example c# | save workbook as html with embedded images asp.net core aspose cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 property | C# generate HTML with inline image data URIs from Excel | Excel to HTML conversion with embedded images using Aspose | Base64 image embedding in Aspose.Cells HTML export | Save workbook as HTML with embedded pictures in .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, optionally adds a picture, configures HtmlSaveOptions to export images as Base64 strings, and saves the result as an HTML file where all worksheet images are embedded directly as data URIs.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put some sample data
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // OPTIONAL: Insert an image to demonstrate base64 embedding
        // Ensure the image file exists at the specified path
        string imagePath = "sample.png";
        if (File.Exists(imagePath))
        {
            // Add the picture to the worksheet (row 2, column 0)
            sheet.Pictures.Add(2, 0, imagePath);
        }

        // Configure HTML save options to embed images as Base64 strings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportImagesAsBase64 = true; // Embed images directly in HTML

        // Save the workbook as an HTML file with embedded images
        workbook.Save("output.html", htmlOptions);
    }
}
