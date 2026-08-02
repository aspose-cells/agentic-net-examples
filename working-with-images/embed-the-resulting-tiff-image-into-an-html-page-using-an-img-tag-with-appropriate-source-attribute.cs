using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF to HTML Demo");

        // Configure image options for TIFF rendering
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,   // Render as TIFF
            OnePagePerSheet = true        // One page per sheet
        };

        // Render the worksheet to a TIFF image in memory
        using (MemoryStream tiffStream = new MemoryStream())
        {
            SheetRender renderer = new SheetRender(worksheet, imgOptions);
            renderer.ToTiff(tiffStream);          // Render to stream
            tiffStream.Position = 0;              // Reset stream position

            // Convert the TIFF bytes to a Base64 string
            string base64Tiff = Convert.ToBase64String(tiffStream.ToArray());

            // Build an HTML page that embeds the TIFF image using a data URI
            string htmlContent = $@"
<html>
<head><title>Embedded TIFF Image</title></head>
<body>
    <h2>Embedded TIFF Image</h2>
    <img src=""data:image/tiff;base64,{base64Tiff}"" alt=""TIFF Image"" />
</body>
</html>";

            // Save the HTML file
            File.WriteAllText("EmbeddedTiff.html", htmlContent);
            Console.WriteLine("HTML file with embedded TIFF created: EmbeddedTiff.html");
        }
    }
}