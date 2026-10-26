// Title: Create a Base64‑encoded PNG data URI from an Excel worksheet using Aspose.Cells for CSS background in C#
// Common Searches: how to convert an Excel worksheet to a Base64 PNG data URI with Aspose.Cells in C# | embed an Excel sheet image as a CSS background using Aspose.Cells rendering | generate a data:image/png;base64 string from a worksheet without writing to disk C# | C# code to render a worksheet to memory stream and use it in CSS background-image | Aspose.Cells one page per sheet option for PNG image generation
// Tags: Aspose.Cells render worksheet to PNG | C# generate Base64 image data URI | Excel worksheet image for CSS background | memory stream PNG rendering with Aspose.Cells | one page per sheet image option

using System;
using System.IO;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders the first worksheet to a PNG image in a memory stream using Aspose.Cells, converts the image bytes to a Base64 string, builds a data:image/png;base64 URI, and outputs a CSS background-image rule that embeds the worksheet image directly.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set image options for PNG rendering (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                Transparent = true
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

            // Use a memory stream to avoid writing to disk
            using (MemoryStream pngStream = new MemoryStream())
            {
                // Render the first page (index 0) of the worksheet to the stream
                sheetRender.ToImage(0, pngStream);

                // Convert the PNG bytes to a Base64 string
                string base64 = Convert.ToBase64String(pngStream.ToArray());

                // Build the data URI for embedding in CSS
                string dataUri = $"data:image/png;base64,{base64}";

                // Example CSS background property using the data URI
                string cssBackground = $"background-image: url('{dataUri}');";

                // Output the CSS string
                Console.WriteLine(cssBackground);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
