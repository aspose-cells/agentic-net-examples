using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffRender
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Render this worksheet to TIFF in a memory stream.");

            // Configure rendering options for TIFF output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                 // Specify TIFF format
                TiffCompression = TiffCompression.CompressionLZW, // Optional compression
                OnePagePerSheet = true                      // Render the whole sheet as one page
            };

            // Initialize the workbook renderer with the workbook and options
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Create a memory stream to hold the TIFF image
            using (MemoryStream tiffStream = new MemoryStream())
            {
                // Render the entire workbook to the memory stream as a TIFF image
                renderer.ToImage(tiffStream);

                // Reset the stream position if it will be read later
                tiffStream.Position = 0;

                // Example: display the size of the generated TIFF data
                Console.WriteLine($"TIFF image generated. Stream length: {tiffStream.Length} bytes");

                // Further processing can be done here using tiffStream
                // For instance, save to a file for verification (optional)
                // File.WriteAllBytes("output.tiff", tiffStream.ToArray());
            }

            // Clean up resources
            renderer.Dispose();
        }
    }
}