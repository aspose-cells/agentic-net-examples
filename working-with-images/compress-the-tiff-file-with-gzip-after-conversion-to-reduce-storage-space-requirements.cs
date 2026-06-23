using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffGzipDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("TIFF Compression with GZip Demo");
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
                    renderer.ToTiff(tiffStream);
                    tiffStream.Position = 0; // Reset stream position for reading

                    // Define the output GZip file path
                    string gzipPath = "Worksheet.tiff.gz";

                    // Ensure the directory for the output file exists
                    string outputDir = Path.GetDirectoryName(gzipPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Compress the TIFF data using GZipStream
                    using (FileStream fileStream = new FileStream(gzipPath, FileMode.Create, FileAccess.Write))
                    using (GZipStream gzipStream = new GZipStream(fileStream, CompressionLevel.Optimal))
                    {
                        tiffStream.CopyTo(gzipStream);
                    }

                    Console.WriteLine($"TIFF rendered and compressed to GZip file: {gzipPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}