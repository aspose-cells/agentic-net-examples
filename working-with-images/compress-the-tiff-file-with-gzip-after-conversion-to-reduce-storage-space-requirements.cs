// Title: Compress a TIFF worksheet with GZip using Aspose.Cells for .NET
// Description: Creates a Workbook, renders the first worksheet to a TIFF stream with optional LZW compression and 300 dpi resolution, then writes the stream into a GZipStream to produce a .tiff.gz file, reducing storage size.
// Keywords: Aspose.Cells | C# | .NET | TIFF compression | GZip | LZW | ImageOrPrintOptions | SheetRender | memory stream | archive Excel as image | reduce file size
// Common Searches: Aspose.Cells render worksheet to TIFF and gzip | C# compress TIFF with GZip | How to create .tiff.gz from Excel using Aspose | LZW TIFF compression with GZip in .NET | Save Excel sheet as compressed TIFF
// Developer Intent: Generate a TIFF image of an Excel worksheet and shrink its footprint by applying GZip compression.
// Use Cases: Long‑term archival of Excel reports as compact TIFF files. | Sending worksheet snapshots over low‑bandwidth networks. | Creating backup images where disk space is limited.
// AI Prompts: Show C# code that uses Aspose.Cells to render a worksheet to a TIFF stream with LZW compression, then compress it with GZip and save as .tiff.gz. | Explain how to adjust ImageOrPrintOptions for resolution and TIFF compression before GZip compression. | Give best‑practice error handling for rendering a worksheet to TIFF and compressing it with GZip in a .NET application.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffGzipExample
{
    // Creates a Workbook, renders the first worksheet to a TIFF stream with optional LZW compression and 300 dpi resolution, then writes the stream into a GZipStream to produce a .tiff.gz file, reducing storage size.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("TIFF GZip Compression Demo");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // Configure image options for TIFF rendering
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Use LZW compression for the TIFF itself (optional)
                    TiffCompression = TiffCompression.CompressionLZW,
                    // Set a reasonable resolution
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Render the worksheet to a memory stream as TIFF
                using (MemoryStream tiffStream = new MemoryStream())
                {
                    SheetRender renderer = new SheetRender(sheet, imgOptions);
                    renderer.ToTiff(tiffStream); // Render to TIFF stream

                    // Prepare the output file for the GZip-compressed TIFF
                    using (FileStream outputFile = new FileStream("Worksheet.tiff.gz", FileMode.Create, FileAccess.Write))
                    using (GZipStream gzip = new GZipStream(outputFile, CompressionLevel.Optimal))
                    {
                        // Reset the position of the TIFF stream before copying
                        tiffStream.Position = 0;
                        tiffStream.CopyTo(gzip);
                    }
                }

                Console.WriteLine("Worksheet rendered to TIFF and compressed with GZip successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
