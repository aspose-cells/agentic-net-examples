// Title: Cache rendered worksheet PNG images in memory with Aspose.Cells for .NET to reuse across requests
// Common Searches: c# cache worksheet png bytes using Aspose.Cells for repeated rendering | store Aspose.Cells rendered worksheet image in session memory | how to avoid re‑rendering Excel sheet to PNG with Aspose.Cells | in‑memory cache for SheetRender PNG output in ASP.NET | reuse Aspose.Cells worksheet PNG across multiple HTTP requests
// Tags: in‑memory cache for Aspose.Cells worksheet PNG | SheetRender PNG rendering with ImageOrPrintOptions | static dictionary cache keyed by workbook path and sheet index | write cached PNG byte array to file using Aspose.Cells | exception handling for worksheet image rendering in C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The WorksheetImageCache class demonstrates how to render an Excel worksheet to a PNG image with Aspose.Cells, store the resulting byte array in a static in‑memory dictionary keyed by the workbook's full path and sheet index, and retrieve it on subsequent calls to avoid redundant rendering. It also provides a helper method to write the cached PNG bytes to disk, includes path validation, directory creation, and robust exception handling for reliable reuse across multiple requests within a .NET session.
public static class WorksheetImageCache
{
    // Simple in‑memory cache for rendered PNG bytes.
    private static readonly Dictionary<string, byte[]> _cache = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

    /// <param name="workbookPath">Full path to the Excel workbook.</param>
    /// <param name="sheetIndex">Zero‑based index of the worksheet to render.</param>
    /// <returns>PNG image bytes.</returns>
    public static byte[] GetWorksheetPng(string workbookPath, int sheetIndex = 0)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
            throw new ArgumentException("Workbook path must be provided.", nameof(workbookPath));

        // Ensure the file exists before attempting to load it.
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"The workbook file was not found: {workbookPath}");

        // Build a unique cache key per workbook and sheet.
        string cacheKey = $"WorksheetImage_{Path.GetFullPath(workbookPath)}_{sheetIndex}";

        // Return cached image if available.
        if (_cache.TryGetValue(cacheKey, out byte[] cachedBytes) && cachedBytes?.Length > 0)
            return cachedBytes;

        try
        {
            // Load the workbook using Aspose.Cells.
            var workbook = new Workbook(workbookPath);

            // Validate sheet index.
            if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Sheet index is out of range.");

            var worksheet = workbook.Worksheets[sheetIndex];

            // Configure image rendering options. Default format is PNG, so we omit explicit ImageFormat.
            var imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to an image.
            var sheetRender = new SheetRender(worksheet, imgOptions);
            using (var ms = new MemoryStream())
            {
                sheetRender.ToImage(0, ms);
                byte[] pngBytes = ms.ToArray();

                // Cache the rendered PNG for future calls.
                _cache[cacheKey] = pngBytes;

                return pngBytes;
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context while preserving the original stack trace.
            throw new InvalidOperationException($"Failed to render worksheet index {sheetIndex} from workbook '{workbookPath}'.", ex);
        }
    }

    /// <param name="workbookPath">Full path to the Excel workbook.</param>
    /// <param name="outputPath">File path where the PNG will be saved.</param>
    /// <param name="sheetIndex">Zero‑based index of the worksheet to render.</param>
    public static void WriteWorksheetPngToFile(string workbookPath, string outputPath, int sheetIndex = 0)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path must be provided.", nameof(outputPath));

        byte[] pngBytes = GetWorksheetPng(workbookPath, sheetIndex);

        try
        {
            // Ensure the directory exists.
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            File.WriteAllBytes(outputPath, pngBytes);
        }
        catch (Exception ex)
        {
            throw new IOException($"Failed to write PNG to '{outputPath}'.", ex);
        }
    }
}

// Simple entry point to satisfy the project build.
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage (adjust paths as needed). Wrapped in try‑catch to avoid unhandled exceptions.
        try
        {
            string workbookPath = "sample.xlsx"; // Replace with an existing workbook path.
            string outputPath = "output.png";

            // Render first worksheet to PNG and save to file.
            WorksheetImageCache.WriteWorksheetPngToFile(workbookPath, outputPath);
            Console.WriteLine($"Worksheet rendered and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
