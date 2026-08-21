// Title: Export Excel to MHTML with UTF‑8 Encoding & Embedded Base64 Resources using Aspose.Cells for .NET
// Description: Shows how to build a Workbook, add sample data, configure HtmlSaveOptions for MHTML, set UTF‑8 encoding, embed images as Base64, and save the result as an .mht file with Aspose.Cells in C#.
// Keywords: Aspose.Cells | MHTML export | UTF-8 encoding | Base64 image embedding | HtmlSaveOptions | C# Excel to MHTML | embedded resources | .NET conversion | save as .mht | Excel web archive
// Common Searches: Aspose.Cells export to MHTML | C# save Excel as .mht UTF-8 | embed images in MHTML using Aspose | HtmlSaveOptions MHtml example | convert workbook to MHTML with base64 images | set encoding for MHTML in Aspose.Cells
// Developer Intent: Generate an MHTML document from an Excel workbook with UTF‑8 encoding and all assets embedded as Base64 using Aspose.Cells for .NET.
// Use Cases: Create a single‑file MHTML report that can be attached to email without external assets. | Provide a web‑ready preview of a spreadsheet, preserving charts and pictures within the file. | Automate archival of Excel workbooks as self‑contained MHTML files while maintaining Unicode characters. | Integrate MHTML generation into a .NET service that delivers spreadsheet content to browsers.
// AI Prompts: Modify the example to include a chart image in the MHTML output with Aspose.Cells. | Show how to write the MHTML to a MemoryStream instead of a physical file while keeping UTF‑8 and Base64 resources. | Explain how to apply custom CSS styles to the generated MHTML using HtmlSaveOptions.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMhtmlExample
{
    // Shows how to build a Workbook, add sample data, configure HtmlSaveOptions for MHTML, set UTF‑8 encoding, embed images as Base64, and save the result as an .mht file with Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the default constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, MHTML!");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Optionally add an image to demonstrate embedded resources
            // (replace "example.jpg" with a valid image path if needed)
            // sheet.Pictures.Add(1, 1, "example.jpg");

            // Create HTML save options for MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);

            // Set UTF‑8 encoding
            saveOptions.Encoding = Encoding.UTF8;

            // Embed images (and other resources) as Base64 strings
            saveOptions.ExportImagesAsBase64 = true;

            // Save the workbook as an MHTML file with the specified options
            string outputPath = "output.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"MHTML file saved to: {outputPath}");
        }
    }
}
