// Title: Cache rendered worksheet PNG images in memory using Aspose.Cells for .NET
// AI Prompts: Create a static method that renders a specified worksheet to a PNG byte array with Aspose.Cells and stores the result in a ConcurrentDictionary for thread‑safe reuse. | Add validation logic to ensure the workbook file exists and the sheet index is within range before rendering. | Demonstrate retrieving the cached PNG bytes from the in‑memory store and writing them to a file in a simple console program.
// Common Searches: how to store Aspose.Cells worksheet PNG in memory for reuse in C# | c# Aspose.Cells render worksheet to PNG and cache result | thread‑safe in‑memory cache for Excel sheet images Aspose.Cells | reuse rendered worksheet image across multiple web requests .NET | Aspose.Cells SheetRender cache PNG bytes per worksheet
// Tags: Aspose.Cells render worksheet to PNG bytes | thread‑safe in‑memory cache using ConcurrentDictionary | worksheet image caching per workbook path and sheet index | C# Excel sheet PNG generation with Aspose.Cells | reuse rendered Excel worksheet image in .NET

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.Collections.Concurrent;
using System.IO;

// The example provides a static WorksheetImageCache class that validates input, loads a workbook, renders a chosen worksheet to a PNG byte array using Aspose.Cells' SheetRender with ImageOrPrintOptions, and caches the bytes in a thread‑safe ConcurrentDictionary keyed by workbook path and sheet index. Subsequent calls return the cached image, reducing rendering overhead. A minimal console program shows how to retrieve the cached PNG and save it to disk.
public static class WorksheetImageCache
{
    // Thread‑safe in‑memory cache for generated PNG bytes.
    private static readonly ConcurrentDictionary<string, byte[]> _cache = new ConcurrentDictionary<string, byte[]>();

    // Returns PNG bytes for the specified worksheet, using cache when possible.
    public static byte[] GetWorksheetPng(string workbookPath, int sheetIndex)
    {
        if (string.IsNullOrEmpty(workbookPath))
            throw new ArgumentException("Workbook path must be provided.", nameof(workbookPath));

        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("Workbook file not found.", workbookPath);

        // Load workbook inside try to capture any loading issues.
        Workbook workbook;
        try
        {
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to load workbook '{workbookPath}'.", ex);
        }

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Sheet index is out of range.");

        string cacheKey = $"WorksheetPng_{workbookPath}_{sheetIndex}";

        // Return cached image if available.
        if (_cache.TryGetValue(cacheKey, out var cachedBytes))
            return cachedBytes;

        try
        {
            // Configure rendering options for PNG output.
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // Default format is PNG; explicit setting omitted to avoid missing enum issue.
                OnePagePerSheet = true,
                Transparent = false
            };

            // Render the specified worksheet.
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[sheetIndex], renderOptions);
            using (MemoryStream ms = new MemoryStream())
            {
                sheetRender.ToImage(0, ms); // Page index 0 because OnePagePerSheet = true.
                byte[] pngBytes = ms.ToArray();

                // Cache the result for future calls.
                _cache[cacheKey] = pngBytes;
                return pngBytes;
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception with a clearer message.
            throw new InvalidOperationException("Failed to generate PNG image from worksheet.", ex);
        }
    }
}

// Minimal console entry point to satisfy the project build.
public class Program
{
    public static void Main()
    {
        // Example usage – adjust the path and sheet index as needed.
        string workbookPath = "sample.xlsx";
        try
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Sample workbook not found.", workbookPath);

            byte[] png = WorksheetImageCache.GetWorksheetPng(workbookPath, 0);
            File.WriteAllBytes("sheet0.png", png);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
