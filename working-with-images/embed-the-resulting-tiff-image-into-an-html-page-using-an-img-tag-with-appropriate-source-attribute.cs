using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF to HTML Demo");

        // Configure rendering options for TIFF output
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            OnePagePerSheet = true,
            ImageType = ImageType.Tiff
        };

        // Render the worksheet to a TIFF image stored in a memory stream
        SheetRender sheetRenderer = new SheetRender(worksheet, renderOptions);
        using (MemoryStream tiffStream = new MemoryStream())
        {
            sheetRenderer.ToTiff(tiffStream); // Render to TIFF stream

            // Convert the TIFF bytes to a Base64 string
            string base64Tiff = Convert.ToBase64String(tiffStream.ToArray());

            // Build an HTML document that embeds the TIFF image using a data URI
            string htmlContent = $@"
<html>
<head><title>Embedded TIFF Image</title></head>
<body>
    <h2>Embedded TIFF Image</h2>
    <img src=""data:image/tiff;base64,{base64Tiff}"" alt=""TIFF Image"" />
</body>
</html>";

            // Save the HTML to a file
            File.WriteAllText("output.html", htmlContent);
        }

        Console.WriteLine("HTML file with embedded TIFF image has been created.");
    }
}