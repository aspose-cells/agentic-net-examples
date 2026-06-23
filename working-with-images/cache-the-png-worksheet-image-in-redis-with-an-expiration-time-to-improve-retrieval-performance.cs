using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

namespace AsposeCellsRedisCacheDemo
{
    class Program
    {
        // Path to store the generated PNG (used instead of Redis for this demo)
        private const string OutputImagePath = "worksheet_image.png";

        static async Task Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and add some data
                var workbook = new Workbook(); // creates a new workbook
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Cached PNG Image Demo");
                worksheet.Cells["A2"].PutValue(DateTime.Now.ToString());

                // 2. Render the first worksheet to a PNG image in memory
                var renderOptions = new ImageOrPrintOptions
                {
                    // PNG is the default format; no need to set ImageFormat explicitly
                    OnePagePerSheet = true
                };

                var sheetRender = new SheetRender(worksheet, renderOptions);
                byte[] pngBytes;
                using (var ms = new MemoryStream())
                {
                    // Render page 0 (the only page because OnePagePerSheet = true)
                    sheetRender.ToImage(0, ms);
                    pngBytes = ms.ToArray();
                }

                // 3. Store the PNG bytes to a file (simulating a cache store)
                await File.WriteAllBytesAsync(OutputImagePath, pngBytes);
                Console.WriteLine($"Worksheet image saved to '{OutputImagePath}'. Size: {pngBytes.Length} bytes.");

                // Optional: Verify the saved file exists and read its size
                if (File.Exists(OutputImagePath))
                {
                    var fileInfo = new FileInfo(OutputImagePath);
                    Console.WriteLine($"Verification succeeded. File size on disk: {fileInfo.Length} bytes.");
                }
                else
                {
                    Console.WriteLine("Verification failed: file not found.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}