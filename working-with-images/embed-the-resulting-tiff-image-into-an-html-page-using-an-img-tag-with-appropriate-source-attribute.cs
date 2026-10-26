// Title: Generate an HTML page with an embedded PNG image of an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to render an Excel worksheet to a PNG stream and embed it in HTML with Aspose.Cells C# | Aspose.Cells C# create base64 PNG from worksheet and display in <img> tag | C# convert Excel sheet to image and generate HTML file with embedded image using Aspose.Cells | save workbook as PNG in memory and embed as data URI in HTML using Aspose.Cells | Aspose.Cells render sheet to image and write HTML file with base64 image C#
// Tags: Aspose.Cells render worksheet to PNG stream | C# embed base64 image in HTML img tag | SheetRender export Excel sheet as image memory stream | ImageOrPrintOptions OnePagePerSheet configuration | create HTML file with embedded workbook image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a Workbook, adds sample data, uses ImageOrPrintOptions and SheetRender to render the first worksheet to a PNG image stored in a memory stream, converts the image bytes to a Base64 string, builds an HTML document containing an <img> tag with a data:image/png;base64 source, and writes the HTML to a file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["A2"].PutValue(12345);

            // Set image export options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to an image stored in a memory stream
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            using (MemoryStream imageStream = new MemoryStream())
            {
                renderer.ToImage(0, imageStream);
                byte[] imageBytes = imageStream.ToArray();

                // Convert image bytes to Base64 for embedding
                string base64Image = Convert.ToBase64String(imageBytes);

                // Build HTML with an <img> tag referencing the Base64 PNG data
                string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Workbook as Image</title>
</head>
<body>
    <img src=""data:image/png;base64,{base64Image}"" alt=""Workbook Image"" />
</body>
</html>";

                // Determine output path and ensure directory exists
                string outputPath = "WorkbookImage.html";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the HTML to a file
                File.WriteAllText(outputPath, htmlContent);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
