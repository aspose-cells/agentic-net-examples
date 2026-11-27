// Title: Cache rendered worksheet PNG in Redis with expiration using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells render worksheet to PNG and store in Redis with TTL | How to cache Excel sheet image in Redis using Aspose.Cells .NET | Retrieve cached PNG of Excel worksheet from Redis in a .NET application | Set expiration for worksheet image cached in Redis with Aspose.Cells | Aspose.Cells SheetRender cache image bytes in Redis example
// Tags: Aspose.Cells SheetRender PNG output | Redis TTL caching for byte[] images | C# cache rendered Excel worksheet image | Excel to PNG rendering with Aspose.Cells | In‑memory Redis‑like cache implementation .NET

using System;
using System.Collections.Concurrent;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsRedisCache
{
    // Simple in‑memory cache that mimics Redis behaviour for this example
    // The example loads an Excel workbook, uses Aspose.Cells SheetRender to convert the first worksheet page to a PNG byte array, stores the bytes in a Redis‑style cache with a 10‑minute TTL, and demonstrates retrieving the cached image for reuse, writing both the rendered and cached files to disk.
    internal static class SimpleCache
    {
        private static readonly ConcurrentDictionary<string, CacheItem> _cache = new();

        private class CacheItem
        {
            public byte[] Data { get; }
            public DateTime Expiration { get; }

            public CacheItem(byte[] data, DateTime expiration)
            {
                Data = data;
                Expiration = expiration;
            }
        }

        public static void Set(string key, byte[] data, TimeSpan ttl)
        {
            var expiration = DateTime.UtcNow.Add(ttl);
            _cache[key] = new CacheItem(data, expiration);
        }

        public static byte[] Get(string key)
        {
            if (_cache.TryGetValue(key, out var item))
            {
                if (DateTime.UtcNow <= item.Expiration)
                {
                    return item.Data;
                }
                // expired – remove
                _cache.TryRemove(key, out _);
            }
            return null;
        }
    }

    class Program
    {
        // Configuration constants – adjust as needed
        private const string WorkbookPath = "input.xlsx";            // path to the source Excel file
        private const string CacheKeyPrefix = "worksheet_image";
        private static readonly TimeSpan CacheExpiration = TimeSpan.FromMinutes(10);

        static void Main(string[] args)
        {
            try
            {
                // 1. Load the workbook
                if (!File.Exists(WorkbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {WorkbookPath}");
                    return;
                }

                Workbook workbook = new Workbook(WorkbookPath);
                Worksheet sheet = workbook.Worksheets[0]; // using the first worksheet; adjust index if needed

                // 2. Build a cache key that uniquely identifies the worksheet page (first page only)
                string cacheKey = $"{CacheKeyPrefix}:{sheet.Name}:page0";

                // 3. Try to get cached image bytes
                byte[] cachedImage = SimpleCache.Get(cacheKey);
                if (cachedImage != null && cachedImage.Length > 0)
                {
                    Console.WriteLine("Image retrieved from cache.");
                    File.WriteAllBytes("cached_page0.png", cachedImage);
                }
                else
                {
                    Console.WriteLine("Image not found in cache. Rendering with Aspose.Cells...");

                    // 4. Set up rendering options for PNG
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = Aspose.Cells.Drawing.ImageType.Png,
                        OnePagePerSheet = true, // render the whole sheet as one page
                        HorizontalResolution = 150,
                        VerticalResolution = 150
                    };

                    // 5. Render the first page to a memory stream
                    MemoryStream ms = null;
                    SheetRender renderer = null;
                    try
                    {
                        ms = new MemoryStream();
                        renderer = new SheetRender(sheet, options);

                        // Ensure the requested page exists
                        if (renderer.PageCount == 0)
                        {
                            Console.WriteLine("No pages were generated for the worksheet.");
                            return;
                        }

                        // Render page 0 (first page) to the stream
                        renderer.ToImage(0, ms);

                        // 6. Get the byte array
                        byte[] imageBytes = ms.ToArray();

                        // 7. Store the image in cache with expiration
                        SimpleCache.Set(cacheKey, imageBytes, CacheExpiration);
                        Console.WriteLine($"Image cached with key '{cacheKey}' (expires in {CacheExpiration}).");

                        // 8. Write the rendered image to disk for verification
                        File.WriteAllBytes("rendered_page0.png", imageBytes);
                    }
                    catch (Exception renderEx)
                    {
                        Console.WriteLine($"Error during rendering: {renderEx.Message}");
                    }
                    finally
                    {
                        // Dispose resources
                        renderer?.Dispose();
                        ms?.Dispose();
                    }
                }

                // Optional: Demonstrate retrieval of the cached image again
                byte[] verifyBytes = SimpleCache.Get(cacheKey);
                if (verifyBytes != null && verifyBytes.Length > 0)
                {
                    Console.WriteLine("Verification read from cache succeeded.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
