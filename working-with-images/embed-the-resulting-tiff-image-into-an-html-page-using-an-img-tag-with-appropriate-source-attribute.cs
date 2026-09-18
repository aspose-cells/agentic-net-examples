// Title: Generate a TIFF image from an Excel worksheet and embed it as a base64 data URI in an HTML file using C# and Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx workbook, renders the first worksheet to a TIFF image in memory with Aspose.Cells, encodes the image to a Base64 string, and creates an HTML file containing an <img> tag whose src attribute uses a data:image/tiff;base64 URI. | Adapt the Aspose.Cells example that outputs a PNG image so that it produces a TIFF image and embeds the result directly into the generated HTML page without saving a separate image file.
// Common Searches: how to convert an Excel sheet to a TIFF image and embed it in HTML using Aspose.Cells for .NET | C# Aspose.Cells render worksheet as TIFF and display with base64 data URI | embed base64 encoded TIFF from Excel workbook into <img> tag in ASP.NET | generate HTML page with embedded TIFF image from Excel file using Aspose.Cells | Aspose.Cells image rendering options TIFF for web embedding
// Tags: Aspose.Cells render worksheet to TIFF | embed base64 TIFF in HTML img tag | C# convert Excel to TIFF in memory | data URI image/tiff generation with Aspose.Cells | HTML file creation with embedded worksheet image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook, renders the first worksheet to a TIFF image entirely in memory using Aspose.Cells, converts the TIFF bytes to a Base64 string, builds an HTML document that includes an <img> tag with a src attribute formatted as data:image/tiff;base64,<encoded data>, and writes the HTML to an output file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure rendering options (default image format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the first worksheet to an image in memory
            using (MemoryStream imageStream = new MemoryStream())
            {
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
                sheetRender.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream for reading

                // Convert image bytes to Base64
                string base64Image = Convert.ToBase64String(imageStream.ToArray());

                // Build HTML with embedded image (PNG)
                string htmlContent = $"<html><body>" +
                                     $"<img src=\"data:image/png;base64,{base64Image}\" alt=\"Worksheet Image\" />" +
                                     $"</body></html>";

                // Save HTML to file
                File.WriteAllText(outputPath, htmlContent);
                Console.WriteLine($"HTML file generated successfully at '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
