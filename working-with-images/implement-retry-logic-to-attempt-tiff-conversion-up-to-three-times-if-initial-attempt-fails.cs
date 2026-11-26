// Title: Implement retry logic for converting an Excel worksheet to multi‑page TIFF with Aspose.Cells in C#
// Common Searches: c# implement retry loop for Aspose.Cells ToTiff method | how to add back‑off when exporting Excel sheet to TIFF with Aspose.Cells | Aspose.Cells SheetRender retry on conversion failure in C# | using IPageSavingCallback while retrying TIFF export in Aspose.Cells
// Tags: Aspose.Cells retry TIFF conversion | SheetRender ToTiff retry logic | ImageOrPrintOptions LZW TIFF C# | IPageSavingCallback logging | incremental backoff Excel to TIFF | multi‑page TIFF export Aspose.Cells

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;          // For ImageType enum
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffRetryDemo
{
    // Simple implementation of IPageSavingCallback to log page processing.
    // The sample loads an Excel workbook, configures ImageOrPrintOptions for LZW‑compressed 300 dpi TIFF output, and uses SheetRender to convert the first worksheet to a multi‑page TIFF file. It wraps the ToTiff call in a retry loop that attempts the conversion up to three times, applying an incremental delay between attempts, and logs page start/end events via a custom IPageSavingCallback. If all attempts fail, the program exits with a non‑zero status code.
    class ConsolePageSavingCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"[Callback] Starting to save page {args.PageIndex + 1} (IsToOutput = {args.IsToOutput})");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"[Callback] Finished saving page {args.PageIndex + 1} (HasMorePages = {args.HasMorePages})");
        }
    }

    class Program
    {
        // Path to the source workbook. Adjust as needed.
        private const string InputFilePath = "input.xlsx";

        // Desired output TIFF file.
        private const string OutputFilePath = "output.tiff";

        // Maximum number of conversion attempts.
        private const int MaxAttempts = 3;

        static void Main(string[] args)
        {
            // Validate input file existence.
            if (!File.Exists(InputFilePath))
            {
                Console.WriteLine($"Error: Input file \"{InputFilePath}\" does not exist.");
                return;
            }

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(InputFilePath);
                Console.WriteLine($"Loaded workbook \"{InputFilePath}\" successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Use the first worksheet for conversion (adjust if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options for TIFF.
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                 // Set output format to TIFF
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                PageSavingCallback = new ConsolePageSavingCallback()
            };

            // Create the renderer.
            SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

            // Retry loop.
            bool success = false;
            for (int attempt = 1; attempt <= MaxAttempts; attempt++)
            {
                try
                {
                    Console.WriteLine($"Attempt {attempt} of {MaxAttempts}: Converting worksheet to TIFF...");

                    // Render all pages of the worksheet to a single multi‑page TIFF file.
                    sheetRender.ToTiff(OutputFilePath);

                    Console.WriteLine($"TIFF conversion succeeded. Output saved to \"{OutputFilePath}\".");
                    success = true;
                    break; // Exit the retry loop.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");

                    if (attempt == MaxAttempts)
                    {
                        Console.WriteLine("All retry attempts exhausted. Conversion failed.");
                    }
                    else
                    {
                        // Optional back‑off before next attempt.
                        int delayMs = 1000 * attempt; // 1s, 2s, 3s...
                        Console.WriteLine($"Waiting {delayMs} ms before next attempt...");
                        Thread.Sleep(delayMs);
                    }
                }
            }

            if (!success)
            {
                // Indicate failure via exit code.
                Environment.ExitCode = 1;
            }
        }
    }
}
