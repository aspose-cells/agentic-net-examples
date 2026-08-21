// Title: C# – Generate a Base64‑encoded TIFF thumbnail from an Aspose.Cells worksheet and embed it in JSON for API responses
// Description: Creates a Workbook, adds optional content, renders the first worksheet to a single‑page TIFF using SheetRender, converts the TIFF bytes to a Base64 string, and serializes a JSON object with a "thumbnail" property that can be returned from a .NET Web API without writing any files to disk.
// Keywords: Aspose.Cells C# | SheetRender TIFF | Base64 thumbnail | JSON API response | Excel preview image | memory stream rendering | REST service thumbnail | Aspose.Cells example GitHub | C# image to Base64 | Excel to JSON metadata
// Common Searches: Aspose.Cells generate TIFF thumbnail C# | convert worksheet image to Base64 string | return Excel preview as JSON in .NET | SheetRender ToTiff memory stream example | embed Base64 image in API response | C# create Excel thumbnail for web UI
// Developer Intent: Produce a Base64‑encoded TIFF preview of an Excel worksheet and include it in a JSON payload for a web API.
// Use Cases: Provide a lightweight preview of uploaded Excel files in a document‑management portal. | Send a Base64 thumbnail to a JavaScript front‑end for instant display without separate image files. | Cache JSON metadata with the thumbnail to avoid re‑rendering the worksheet on each request.
// AI Prompts: Write C# code that uses Aspose.Cells to render a worksheet to a PNG thumbnail, encode it to Base64, and add it to a JSON API response. | Show how to extend the JSON output with worksheet name, row count, and column count alongside the Base64 TIFF thumbnail. | Create a unit test that verifies the Base64 string generated from SheetRender.ToTiff can be decoded back to a valid TIFF image and matches an expected size.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsThumbnailJson
{
    // Creates a Workbook, adds optional content, renders the first worksheet to a single‑page TIFF using SheetRender, converts the TIFF bytes to a Base64 string, and serializes a JSON object with a "thumbnail" property that can be returned from a .NET Web API without writing any files to disk.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample content (optional, just to have visible data)
            worksheet.Cells["A1"].PutValue("Thumbnail Example");

            // Configure rendering options (single page per sheet)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Initialize SheetRender with the worksheet and options
            SheetRender sheetRenderer = new SheetRender(worksheet, renderOptions);

            // Render the worksheet to a TIFF image in a memory stream (rule: ToTiff(Stream))
            using (MemoryStream tiffStream = new MemoryStream())
            {
                sheetRenderer.ToTiff(tiffStream);

                // Get the TIFF bytes from the stream
                byte[] tiffBytes = tiffStream.ToArray();

                // Convert the TIFF bytes to a Base64 string
                string base64Thumbnail = Convert.ToBase64String(tiffBytes);

                // Build JSON metadata containing the Base64 thumbnail
                var metadata = new
                {
                    thumbnail = base64Thumbnail
                };

                // Serialize the metadata to JSON
                string json = JsonSerializer.Serialize(metadata);

                // Output the JSON (could be returned from an API endpoint)
                Console.WriteLine(json);
            }
        }
    }
}
