// Title: C# – Add 3‑Attempt Retry Logic to Aspose.Cells TIFF Conversion
// Description: Demonstrates how to convert an Aspose.Cells worksheet to a TIFF file with LZW compression while automatically retrying up to three times if the conversion fails. The sample creates a workbook, sets ImageOrPrintOptions, ensures the output folder exists, writes the TIFF via a FileStream, logs each attempt, and re‑throws the exception after the final failure.
// Keywords: Aspose.Cells | C# TIFF conversion | retry logic | SheetRender | ImageOrPrintOptions | LZW compression | file stream | exception handling | ToTiff retry | worksheet to TIFF
// Common Searches: Aspose.Cells retry TIFF conversion C# | C# retry loop for SheetRender ToTiff | How to handle TIFF conversion failures Aspose.Cells | C# example of retrying Aspose.Cells image rendering | Aspose.Cells ToTiff error handling
// Developer Intent: Implement a three‑attempt retry mechanism for Aspose.Cells TIFF conversion in C#.
// Use Cases: Automatically retry worksheet‑to‑TIFF conversion when transient I/O errors occur. | Create the output directory on each attempt to avoid path‑not‑found exceptions. | Log each conversion attempt and surface the error only after all retries fail.
// AI Prompts: Generate C# code that wraps SheetRender.ToTiff in a configurable retry loop with optional exponential backoff. | Refactor the demo into a reusable method that accepts a workbook, worksheet, output path, and max retry count. | Explain how to catch specific exceptions (e.g., IOException) during TIFF conversion and retry only those while letting other exceptions propagate.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to convert an Aspose.Cells worksheet to a TIFF file with LZW compression while automatically retrying up to three times if the conversion fails. The sample creates a workbook, sets ImageOrPrintOptions, ensures the output folder exists, writes the TIFF via a FileStream, logs each attempt, and re‑throws the exception after the final failure.
    public class TiffConversionWithRetryDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample content
            worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF Conversion with Retry");

            // Configure image rendering options for TIFF
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // No need to set ImageFormat when using ToTiff; it is implied
                OnePagePerSheet = true,
                TiffCompression = TiffCompression.CompressionLZW
            };

            const int maxRetryTimes = 3; // maximum number of attempts
            int attempt = 0;
            bool succeeded = false;

            while (attempt < maxRetryTimes && !succeeded)
            {
                attempt++;
                try
                {
                    // Ensure the output directory exists
                    string outputPath = "output.tiff";
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (string.IsNullOrEmpty(outputDir))
                    {
                        outputDir = Directory.GetCurrentDirectory();
                    }

                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Render the worksheet to a TIFF file using a file stream
                    using (FileStream tiffStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        SheetRender renderer = new SheetRender(worksheet, options);
                        renderer.ToTiff(tiffStream);
                    }

                    succeeded = true;
                    Console.WriteLine($"TIFF conversion succeeded on attempt {attempt}.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");
                    if (attempt == maxRetryTimes)
                    {
                        // Re‑throw after the final attempt so the caller can handle the failure
                        throw;
                    }
                }
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                TiffConversionWithRetryDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
