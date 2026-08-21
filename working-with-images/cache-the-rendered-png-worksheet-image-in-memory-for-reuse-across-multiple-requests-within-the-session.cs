// Title: Cache Worksheet PNG in Memory with Aspose.Cells for .NET – Reuse Across Requests
// Description: Demonstrates a static in‑memory cache for a PNG rendered from an Aspose.Cells worksheet using SheetRender and ImageOrPrintOptions. The image is generated once per session and can be returned as a byte array or written to any stream, eliminating repeated rendering in web or desktop apps.
// Keywords: Aspose.Cells | C# | .NET | worksheet PNG cache | in‑memory image caching | SheetRender | ImageOrPrintOptions | ASP.NET Core image reuse | performance optimization | file download | GitHub example
// Common Searches: Aspose.Cells cache rendered PNG in memory | reuse worksheet image without re‑rendering .NET | store Aspose.Cells PNG for multiple HTTP responses | in‑memory caching of worksheet images in ASP.NET | how to avoid duplicate rendering of Excel sheet image
// Developer Intent: Store a worksheet PNG in a static byte array so it can be served repeatedly without re‑rendering.
// Use Cases: Return the cached PNG from an ASP.NET Core controller as a FileResult for fast downloads. | Generate a thumbnail once and embed it in PDF reports or email messages multiple times. | Write the cached image to a file, response stream, or cloud storage on demand. | Improve performance of dashboards that display the same worksheet snapshot to many users.
// AI Prompts: Show C# code that injects WorksheetImageCache into an ASP.NET Core MVC action returning a FileResult. | Add thread‑safe locking to the static cache to prevent race conditions in a multi‑threaded web app. | Explain how to invalidate or refresh the cached PNG when the workbook data changes. | Provide a GitHub‑style README section describing how to integrate this cache into a web API project.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates a static in‑memory cache for a PNG rendered from an Aspose.Cells worksheet using SheetRender and ImageOrPrintOptions. The image is generated once per session and can be returned as a byte array or written to any stream, eliminating repeated rendering in web or desktop apps.
public static class WorksheetImageCache
{
    // In-memory cache for the generated PNG image bytes
    private static byte[]? _cachedImage;

    /// <returns>Byte array containing the PNG image.</returns>
    public static byte[] GetWorksheetPng()
    {
        try
        {
            // Return cached image if it already exists
            if (_cachedImage != null && _cachedImage.Length > 0)
                return _cachedImage;

            // Create a new workbook and populate it with sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells cached PNG demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set rendering options for PNG output (default format is PNG)
            var renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to a memory stream
            var renderer = new SheetRender(sheet, renderOptions);
            using (var ms = new MemoryStream())
            {
                renderer.ToImage(0, ms); // Render first page
                _cachedImage = ms.ToArray();
                return _cachedImage;
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception in a more descriptive one
            throw new InvalidOperationException("Failed to generate worksheet PNG.", ex);
        }
    }

    /// <param name="outputStream">The stream to which the PNG data will be written.</param>
    public static void WriteImageToStream(Stream outputStream)
    {
        if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));

        var pngData = GetWorksheetPng();
        if (pngData == null || pngData.Length == 0) return;

        try
        {
            outputStream.Write(pngData, 0, pngData.Length);
        }
        catch (Exception ex)
        {
            // Handle stream write errors
            throw new IOException("Failed to write PNG data to the output stream.", ex);
        }
    }
}

public static class Program
{
    // Entry point for the console application
    public static void Main()
    {
        const string outputPath = "worksheet.png";

        try
        {
            // Ensure the directory exists
            var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Write the cached PNG to a file
            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                WorksheetImageCache.WriteImageToStream(fileStream);
            }

            Console.WriteLine($"Worksheet PNG has been saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
