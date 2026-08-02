// Title: Export Excel to MHTML with UTF‑8 encoding and Base64‑embedded images using Aspose.Cells for .NET
// Description: Shows how to build a workbook, insert text and an optional picture, set HtmlSaveOptions for MHTML with UTF‑8 encoding and ExportImagesAsBase64, and save the result as a self‑contained .mht file.
// Keywords: Aspose.Cells MHTML export | C# UTF-8 MHTML | Base64 images in MHTML | HtmlSaveOptions SaveFormat.MHtml | Aspose.Cells .NET conversion | embedded resources MHTML | Excel to .mht
// Common Searches: Aspose.Cells export workbook to MHTML UTF-8 | C# generate .mht file with embedded images | Save Excel as single‑file MHTML using Aspose | How to embed pictures as Base64 in MHTML output | Convert Excel worksheet to self‑contained MHTML
// Developer Intent: Create a single‑file MHTML document from an Excel workbook, preserving Unicode characters and embedding all images as Base64.
// Use Cases: Send a fully formatted Excel report via email as an MHTML attachment that displays correctly without external files. | Archive Excel dashboards as self‑contained web pages for long‑term storage or offline viewing. | Generate printable HTML previews of spreadsheets for web portals where external image links are not allowed.
// AI Prompts: Provide code to add custom meta tags (title, author) to the generated MHTML file with Aspose.Cells. | Show how to insert multiple pictures into a worksheet and ensure each is embedded as Base64 in the MHTML output. | Explain how to customize CSS styles (fonts, colors) when exporting an Excel sheet to MHTML using HtmlSaveOptions.

using System;
using System.Text;
using System.IO;
using Aspose.Cells;

// Shows how to build a workbook, insert text and an optional picture, set HtmlSaveOptions for MHTML with UTF‑8 encoding and ExportImagesAsBase64, and save the result as a self‑contained .mht file.
class GenerateMhtml
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample text
            sheet.Cells["A1"].PutValue("MHTML Export with UTF-8 encoding and embedded resources");

            // Add an image to be embedded if the file exists
            string imagePath = "example.jpg";
            if (File.Exists(imagePath))
            {
                sheet.Pictures.Add(2, 0, imagePath);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping image insertion.");
            }

            // Create HTML save options for MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
            {
                Encoding = Encoding.UTF8,
                ExportImagesAsBase64 = true
            };

            // Save the workbook as an MHTML document
            workbook.Save("output.mht", saveOptions);
            Console.WriteLine("MHTML file saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
