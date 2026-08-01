// Title: In‑Memory Caching of Worksheet PNG Images with Aspose.Cells for .NET
// Description: Demonstrates a static `WorksheetImageCache` that renders the first worksheet to a PNG using `SheetRender` and `ImageOrPrintOptions`, stores the byte array in a dictionary keyed by a SHA‑256 workbook hash, and returns the cached image on subsequent calls. Includes a simple ASP.NET‑style handler and a console entry point for verification.
// Keywords: Aspose.Cells PNG cache | C# worksheet image rendering | in‑memory image cache .NET | SheetRender PNG Aspose | workbook hash cache key | static dictionary caching | ASP.NET Core image preview | performance optimization Excel rendering
// Common Searches: cache rendered worksheet as PNG Aspose.Cells | reuse SheetRender output in C# web app | generate workbook hash for image cache | store Excel sheet PNG in memory for multiple requests | Aspose.Cells image caching example
// Developer Intent: Store the PNG bytes of a rendered worksheet in memory so the same image can be served repeatedly without re‑rendering the workbook.
// Use Cases: Serve a thumbnail of an uploaded Excel file from a Web API without re‑processing the file on each call. | Accelerate a reporting dashboard by caching worksheet previews for frequently accessed workbooks. | Provide fast document previews in a document‑management system where the same workbook is viewed by many users.
// AI Prompts: Generate thread‑safe C# code that adds LRU eviction to the WorksheetImageCache dictionary. | Show how to return the cached PNG from an ASP.NET Core controller as a FileResult with proper content‑type headers. | Explain how to replace the SHA‑256 hash with MD5 for the cache key and discuss the trade‑offs in speed and security.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Demonstrates a static `WorksheetImageCache` that renders the first worksheet to a PNG using `SheetRender` and `ImageOrPrintOptions`, stores the byte array in a dictionary keyed by a SHA‑256 workbook hash, and returns the cached image on subsequent calls. Includes a simple ASP.NET‑style handler and a console entry point for verification.
public static class WorksheetImageCache
{
    // Simple in‑memory cache keyed by a workbook identifier (e.g., hash of its content)
    private static readonly Dictionary<string, byte[]> _cache = new Dictionary<string, byte[]>();

    // Returns the PNG image bytes of the first worksheet.
    // If the image has been rendered before, the cached bytes are returned.
    public static byte[] GetWorksheetPngImage(Workbook workbook)
    {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        // Create a deterministic key for the workbook.
        string cacheKey = GetWorkbookCacheKey(workbook);

        // Return cached image if it exists.
        if (_cache.TryGetValue(cacheKey, out byte[] cachedBytes))
        {
            return cachedBytes;
        }

        // Render the first worksheet to a PNG image.
        Worksheet worksheet = workbook.Worksheets[0];

        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true
        };

        SheetRender sheetRender = new SheetRender(worksheet, options);

        using (MemoryStream ms = new MemoryStream())
        {
            // Render page 0 (the only page because OnePagePerSheet = true) to the memory stream.
            sheetRender.ToImage(0, ms);

            // Ensure the stream position is at the beginning before reading.
            ms.Position = 0;
            byte[] imageBytes = ms.ToArray();

            // Cache the rendered bytes for future requests.
            _cache[cacheKey] = imageBytes;

            return imageBytes;
        }
    }

    // Helper to generate a cache key for a workbook.
    // This implementation creates a hash of the workbook's binary content.
    private static string GetWorkbookCacheKey(Workbook workbook)
    {
        using (MemoryStream ms = new MemoryStream())
        {
            // Save the workbook to a memory stream (no file I/O).
            workbook.Save(ms, SaveFormat.Xlsx);
            byte[] data = ms.ToArray();

            // Compute a simple hash (e.g., SHA256) to use as the key.
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(data);
                return Convert.ToBase64String(hash);
            }
        }
    }
}

// Example usage within a request‑handling method (e.g., ASP.NET controller action)
public class ExampleHandler
{
    public byte[] HandleRequest()
    {
        // Load or create the workbook.
        Workbook workbook = new Workbook(); // For demonstration; replace with file load if needed.

        // Populate some data (only for demonstration; in real scenarios the workbook may already contain data).
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Cached PNG Example");

        // Retrieve the cached PNG image bytes.
        return WorksheetImageCache.GetWorksheetPngImage(workbook);
    }
}

// Entry point required for console execution.
public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            ExampleHandler handler = new ExampleHandler();
            byte[] pngImage = handler.HandleRequest();

            // Write the PNG to a file for verification.
            string outputPath = "output.png";
            File.WriteAllBytes(outputPath, pngImage);
            Console.WriteLine($"PNG image written to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
