// Title: Convert an Excel workbook to a single HTML file with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions to embed all worksheet images as Base64 data URIs, and saves the result to an HTML file. | Show how to use a MemoryStream with Aspose.Cells to produce HTML output that contains inline Base64 images. | Adapt the example to export only the active worksheet while still embedding images as Base64 data URIs.
// Common Searches: Aspose.Cells C# export workbook to HTML with inline Base64 images | How to embed Excel chart images in HTML using Aspose.Cells .NET | Save Excel as HTML without external image files in C# | HtmlSaveOptions ExportImagesAsBase64 example code | Convert Excel to single HTML page with embedded pictures Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | excel to html base64 image embedding | c# memory stream html generation Aspose.Cells | inline data uri images in html Aspose.Cells | single html output from workbook .net

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, sets HtmlSaveOptions.ExportImagesAsBase64 = true, saves the workbook as HTML to a MemoryStream, reads the HTML string, and writes it to output.html, embedding all worksheet images as Base64 data URIs.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to embed images as Base64 data URIs
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // This flag tells Aspose.Cells to convert all images to Base64 strings
            ExportImagesAsBase64 = true,

            // Export all worksheets (set to true if you only need the active sheet)
            ExportActiveWorksheetOnly = false
        };

        // Save the workbook to a memory stream using the HTML options
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, htmlOptions);
            ms.Position = 0;

            // Read the generated HTML content from the memory stream
            string htmlContent = new StreamReader(ms).ReadToEnd();

            // Write the HTML (with embedded Base64 images) to an output file
            File.WriteAllText("output.html", htmlContent);
        }
    }
}
