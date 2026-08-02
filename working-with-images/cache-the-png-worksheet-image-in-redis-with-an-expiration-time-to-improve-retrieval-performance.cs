// Title: Cache a worksheet PNG in Redis with TTL using Aspose.Cells for .NET
// Description: Demonstrates how to render an Excel worksheet to a PNG with Aspose.Cells, store the image bytes in Redis, and apply an expiration time so subsequent requests can fetch the cached image quickly while keeping the data fresh.
// Keywords: Aspose.Cells PNG rendering | Redis cache .NET | Excel worksheet image caching | TTL Redis image | C# Aspose.Cells Redis example | image rendering performance | Excel to PNG Redis | cache expiration strategy
// Common Searches: store Aspose.Cells worksheet image in Redis | cache Excel sheet PNG with expiration C# | Aspose.Cells render worksheet to PNG and cache | Redis TTL for Excel image bytes | how to improve worksheet image retrieval performance
// Developer Intent: Persist a rendered worksheet PNG in Redis and retrieve it efficiently until the configured TTL expires.
// Use Cases: Serve the same worksheet image to many web users without re‑rendering on each request. | Reduce CPU load on a reporting server by caching Excel‑to‑PNG conversions for a limited time. | Provide fast thumbnail previews in a dashboard while automatically refreshing after the cache period. | Implement a scalable image cache for a multi‑instance ASP.NET Core API using Redis.
// AI Prompts: Generate C# code that replaces the in‑memory dictionary with StackExchange.Redis to cache worksheet PNGs and set a 30‑minute expiration. | Show how to create a Redis key that uniquely identifies a worksheet based on file path and sheet index. | Explain fallback logic when Redis is unavailable, rendering the image on‑the‑fly with Aspose.Cells. | Provide an async version of the Redis caching routine for high‑throughput web APIs.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to render an Excel worksheet to a PNG with Aspose.Cells, store the image bytes in Redis, and apply an expiration time so subsequent requests can fetch the cached image quickly while keeping the data fresh.
public class WorksheetImageCache
{
    private readonly TimeSpan _cacheExpiration;
    private readonly Dictionary<string, (byte[] Data, DateTime Expiry)> _cache;

    public WorksheetImageCache(TimeSpan cacheExpiration)
    {
        _cacheExpiration = cacheExpiration;
        _cache = new Dictionary<string, (byte[] Data, DateTime Expiry)>();
    }

    public byte[] GetWorksheetImage(string workbookPath, int worksheetIndex = 0)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
            throw new ArgumentException("Workbook path must be provided.", nameof(workbookPath));

        // Build a unique cache key for the worksheet image
        string cacheKey = $"WorksheetImage:{workbookPath}:{worksheetIndex}";

        // Check in‑memory cache
        if (_cache.TryGetValue(cacheKey, out var entry))
        {
            if (DateTime.UtcNow < entry.Expiry)
                return entry.Data; // Return cached image
            else
                _cache.Remove(cacheKey); // Expired
        }

        // Verify the workbook file exists
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);
            Worksheet worksheet = workbook.Worksheets[worksheetIndex];

            // Configure rendering options for PNG output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };

            // Render the worksheet to a memory stream
            SheetRender renderer = new SheetRender(worksheet, options);
            using (MemoryStream ms = new MemoryStream())
            {
                renderer.ToImage(0, ms);
                byte[] imageBytes = ms.ToArray();

                // Store in cache with expiration
                _cache[cacheKey] = (imageBytes, DateTime.UtcNow.Add(_cacheExpiration));

                return imageBytes;
            }
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed; for this example we rethrow
            throw new InvalidOperationException("Failed to render worksheet image.", ex);
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Initialize cache with 30‑minute expiration
            var cache = new WorksheetImageCache(TimeSpan.FromMinutes(30));

            // Path to the source Excel workbook
            string excelPath = "sample.xlsx";

            // Retrieve the PNG image bytes (from cache if available, otherwise render)
            byte[] pngData = cache.GetWorksheetImage(excelPath);

            // Write the image to a file to verify the result
            File.WriteAllBytes("cached_sheet.png", pngData);
            Console.WriteLine("Worksheet image saved to cached_sheet.png");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
