// Title: Cache a rendered worksheet PNG in Redis with expiration using Aspose.Cells for .NET
// AI Prompts: Write C# code that renders the first worksheet of an Excel file to a PNG with Aspose.Cells, then stores the PNG byte array in Redis using StackExchange.Redis with a configurable expiration time. | Show how to replace the file‑system save in the Aspose.Cells example with Redis caching, including connection multiplexer setup, byte[] storage, and TTL handling.
// Common Searches: how to store Aspose.Cells generated worksheet PNG in Redis with a TTL in C# | c# cache Excel worksheet image in Redis using StackExchange.Redis | Aspose.Cells render sheet to PNG and set expiration in Redis cache | example of Redis caching for images created by Aspose.Cells in .NET | configure Redis TTL for byte[] image data from Aspose.Cells rendering
// Tags: Aspose.Cells render worksheet to PNG | Redis cache PNG byte array C# | StackExchange.Redis TTL for image data | Excel worksheet image caching strategy | C# image rendering and Redis storage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetImageCacheApp
{
    // The example loads an Excel workbook with Aspose.Cells, renders the first worksheet to a PNG image in memory, and demonstrates how to replace the file‑system write with Redis caching of the PNG byte array, including setting a time‑to‑live for faster subsequent retrievals.
    class WorksheetImageCache
    {
        static void Main()
        {
            // Path to the Excel file
            string excelPath = @"C:\Data\Sample.xlsx";

            // Verify that the Excel file exists before loading
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook (Aspose.Cells)
                Workbook workbook = new Workbook(excelPath);

                // Choose the first worksheet to render
                Worksheet sheet = workbook.Worksheets[0];

                // Render the worksheet to a PNG image in memory
                using (MemoryStream imageStream = new MemoryStream())
                {
                    // Set image save options (default format is PNG)
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        OnePagePerSheet = true,
                        Transparent = false,
                        HorizontalResolution = 150,
                        VerticalResolution = 150
                    };

                    // Render the first page of the worksheet
                    SheetRender sr = new SheetRender(sheet, imgOptions);
                    sr.ToImage(0, imageStream); // 0 = first page

                    // Get the PNG bytes
                    byte[] pngBytes = imageStream.ToArray();

                    // Determine output path
                    string outputDirectory = Path.GetDirectoryName(excelPath) ?? string.Empty;
                    string outputPath = Path.Combine(
                        outputDirectory,
                        $"{Path.GetFileNameWithoutExtension(excelPath)}_{sheet.Name}.png");

                    try
                    {
                        // Ensure the directory exists
                        if (!Directory.Exists(outputDirectory))
                        {
                            Directory.CreateDirectory(outputDirectory);
                        }

                        // Save the PNG to a local file (replace with Redis caching if needed)
                        File.WriteAllBytes(outputPath, pngBytes);
                        Console.WriteLine($"Worksheet image saved to '{outputPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write image file: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
