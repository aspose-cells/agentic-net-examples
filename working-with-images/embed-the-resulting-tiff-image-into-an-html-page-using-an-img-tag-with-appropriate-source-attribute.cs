// Title: Embed a TIFF image rendered from an Aspose.Cells worksheet into HTML using C#
// Description: Shows how to create a workbook, render the first worksheet to a TIFF image with Aspose.Cells, convert the image to a Base64 data URI, and generate an HTML file that displays the TIFF via an <img> tag.
// Keywords: Aspose.Cells | C# | TIFF rendering | Base64 data URI | embed image in HTML | SheetRender | ToTiff | generate HTML page | workbook to image | data:image/tiff
// Common Searches: Aspose.Cells render worksheet to TIFF C# | embed TIFF in HTML using data URI | C# convert image to Base64 for web page | display Excel sheet as image in browser | Aspose.Cells SheetRender example
// Developer Intent: Create an HTML page that shows a worksheet as an embedded TIFF image.
// Use Cases: Email a spreadsheet preview without attaching separate image files. | Show Excel data in a web dashboard while avoiding extra HTTP requests. | Generate printable HTML reports with the worksheet rendered as a graphic. | Reduce server storage by using data URIs instead of saved image files. | Provide instant thumbnail previews of spreadsheets in a web application.
// AI Prompts: Write C# code that uses Aspose.Cells to render a worksheet to a TIFF stream and embed it in HTML via a data:image/tiff;base64 URI. | Explain how to apply compression settings to the TIFF output and set the HTML title dynamically from worksheet content. | Create a unit test that verifies the generated HTML contains a valid data:image/tiff;base64 string after rendering. | Show how to modify the example to output PNG instead of TIFF and embed it in the same HTML structure. | Provide steps to deploy the generated HTML file to Azure Static Web Apps for public access.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffToHtml
{
    // Shows how to create a workbook, render the first worksheet to a TIFF image with Aspose.Cells, convert the image to a Base64 data URI, and generate an HTML file that displays the TIFF via an <img> tag.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add some sample data to the worksheet
                worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF to HTML Demo");
                worksheet.Cells["A2"].PutValue(DateTime.Now);

                // Configure image rendering options for TIFF output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Render the whole sheet on one page
                    OnePagePerSheet = true
                };

                // Create a SheetRender instance with the worksheet and options
                SheetRender renderer = new SheetRender(worksheet, imgOptions);

                // Render the worksheet to a memory stream in TIFF format
                using (MemoryStream tiffStream = new MemoryStream())
                {
                    renderer.ToTiff(tiffStream); // Render to TIFF
                    byte[] tiffBytes = tiffStream.ToArray();

                    // Convert the TIFF bytes to a Base64 string
                    string base64Tiff = Convert.ToBase64String(tiffBytes);

                    // Build an HTML page that embeds the TIFF image using a data URI
                    string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Worksheet as TIFF</title>
</head>
<body>
    <h2>Embedded TIFF Image</h2>
    <img src=""data:image/tiff;base64,{base64Tiff}"" alt=""Worksheet TIFF"" />
</body>
</html>";

                    // Ensure the output directory exists
                    string outputPath = "Worksheet.html";
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the HTML file to disk
                    File.WriteAllText(outputPath, htmlContent);
                    Console.WriteLine($"HTML file with embedded TIFF image has been created: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
