// Title: C# – Convert an Aspose.Cells worksheet to PNG and embed as Base64 in JSON
// Description: Demonstrates how to create a workbook, populate cells, render the first worksheet to a PNG image using SheetRender and ImageOrPrintOptions, convert the image stream to a Base64 string, and serialize it into a JSON object printed to the console.
// Keywords: Aspose.Cells PNG export C# | worksheet to image base64 | SheetRender memory stream | ImageOrPrintOptions one page per sheet | serialize image to JSON .NET | C# Excel preview base64
// Common Searches: Aspose.Cells render worksheet as PNG C# | convert Excel sheet to base64 string | C# example: worksheet image to JSON | how to use SheetRender with MemoryStream | base64 encoded Excel preview for API
// Developer Intent: Generate a PNG snapshot of a worksheet and deliver it as a Base64‑encoded value inside a JSON payload.
// Use Cases: Return a worksheet preview in a REST API response for web clients. | Store Excel sheet images as Base64 fields in NoSQL documents for quick retrieval. | Push real‑time worksheet thumbnails over SignalR or WebSocket connections.
// AI Prompts: Write C# code that renders a specific worksheet page to JPEG and returns the Base64 string in a JSON object using Aspose.Cells. | Explain how to adjust ImageOrPrintOptions to change DPI and image format when converting a worksheet to a Base64‑encoded JSON string. | Show how to loop through all worksheets, render each to PNG, and build a JSON array containing worksheet names and their Base64 images.

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate cells, render the first worksheet to a PNG image using SheetRender and ImageOrPrintOptions, convert the image stream to a Base64 string, and serialize it into a JSON object printed to the console.
    public class WorksheetToPngBase64Json
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample data
                worksheet.Cells["A1"].PutValue("Sample");
                worksheet.Cells["B1"].PutValue(123);

                // Set image rendering options (default format is PNG)
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true // Render each sheet as a single page
                };

                // Render the worksheet to a PNG image in memory
                SheetRender sheetRender = new SheetRender(worksheet, options);
                using (MemoryStream imageStream = new MemoryStream())
                {
                    sheetRender.ToImage(0, imageStream);

                    // Convert the rendered image bytes to a Base64 string
                    string base64Image = Convert.ToBase64String(imageStream.ToArray());

                    // Create a simple JSON object containing the Base64 image
                    var jsonObject = new { imageBase64 = base64Image };
                    string json = JsonSerializer.Serialize(jsonObject);

                    // Output the JSON string
                    Console.WriteLine(json);
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetToPngBase64Json.Run();
        }
    }
}
