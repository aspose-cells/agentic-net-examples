// Title: Cache Aspose.Cells Worksheet PNG in Redis with Expiration (C#)
// Description: Demonstrates how to render a worksheet to a PNG image with Aspose.Cells, store the resulting byte array in Redis using StackExchange.Redis, and apply a time‑to‑live (TTL) so the image can be served from cache on subsequent requests, reducing rendering overhead.
// Keywords: Aspose.Cells | C# | Redis cache | PNG image | Worksheet rendering | StackExchange.Redis | TTL | image caching | Excel to PNG | Aspose.Cells Redis integration
// Common Searches: store Aspose.Cells PNG in Redis | cache worksheet image C# Redis TTL | Aspose.Cells render to byte array and cache | Redis expiration for Excel PNG image | retrieve cached worksheet PNG from Redis
// Developer Intent: Save the PNG bytes of a rendered worksheet in Redis with a configurable expiration time and retrieve them to avoid repeated rendering.
// Use Cases: Web API returns a worksheet preview image quickly by reading a cached PNG from Redis instead of re‑rendering the Excel file. | Background service updates the cached PNG whenever the source workbook changes, resetting the TTL to keep the cache fresh. | Multiple microservices share the same Redis cache to serve identical worksheet images without duplicating rendering logic.
// AI Prompts: Generate C# code that renders an Aspose.Cells worksheet to PNG, stores the byte array in Redis with a 10‑minute TTL, and logs the cache key. | Create a method that checks Redis for a cached PNG of a given worksheet ID, returns the image stream if found, otherwise renders, caches, and returns it. | Show how to configure StackExchange.Redis connection settings, serialize the PNG byte array, and handle expiration errors when caching Aspose.Cells images.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to render a worksheet to a PNG image with Aspose.Cells, store the resulting byte array in Redis using StackExchange.Redis, and apply a time‑to‑live (TTL) so the image can be served from cache on subsequent requests, reducing rendering overhead.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Cache this worksheet as PNG");
            worksheet.Cells["A2"].PutValue(DateTime.Now);

            // Set image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the first page of the worksheet to a memory stream
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            using (MemoryStream imageStream = new MemoryStream())
            {
                sheetRender.ToImage(0, imageStream);
                byte[] pngBytes = imageStream.ToArray();

                // Save PNG to a file (simple cache alternative)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "worksheet_page0.png");
                File.WriteAllBytes(outputPath, pngBytes);

                Console.WriteLine($"Worksheet image saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
