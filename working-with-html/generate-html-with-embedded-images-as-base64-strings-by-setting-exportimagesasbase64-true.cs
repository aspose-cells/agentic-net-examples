// Title: Generate HTML from an Excel workbook with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Create an HTML file from a Workbook and embed all worksheet pictures as Base64 strings by configuring HtmlSaveOptions in C#. | Show how to enable HtmlSaveOptions.ExportImagesAsBase64 and save the workbook as an HTML document with embedded image data.
// Common Searches: Aspose.Cells C# export workbook to HTML with images encoded as Base64 | How to embed worksheet pictures as Base64 in HTML output using Aspose.Cells .NET | HtmlSaveOptions ExportImagesAsBase64 example for converting Excel to HTML | Save Excel as HTML with embedded image data using Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | embed worksheet images as Base64 in HTML | C# convert Excel to HTML with embedded images | HTML export with Base64 image encoding Aspose.Cells | save workbook as HTML base64 images

using System;
using Aspose.Cells;

// The sample creates a Workbook, adds data to a cell, optionally inserts a picture, configures HtmlSaveOptions with ExportImagesAsBase64 set to true, and saves the workbook as an HTML file where any worksheet images are embedded directly as Base64 strings.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Optionally insert an image to demonstrate Base64 embedding
        // Replace "sample.png" with a valid image path if needed
        // sheet.Pictures.Add(2, 0, "sample.png");

        // Configure HTML save options to embed images as Base64 strings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;

        // Save the workbook as an HTML file with embedded Base64 images
        workbook.Save("output.html", htmlOptions);
    }
}
