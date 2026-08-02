// Title: Render an Excel worksheet to LZW‑compressed TIFF and GZip it using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply LZW compression via ImageOrPrintOptions, export the first worksheet as a TIFF file, and then shrink the TIFF further by compressing it with GZipStream, all with proper error handling.
// Keywords: Aspose.Cells TIFF export C# | LZW TIFF compression | GZip compress image .NET | reduce Excel export size | C# GZipStream example
// Common Searches: export Aspose.Cells worksheet as compressed TIFF | C# GZip a TIFF generated from Excel | how to shrink TIFF file size with Aspose.Cells | LZW TIFF and GZip compression .NET
// Developer Intent: Create a TIFF image from a workbook with LZW compression and then GZip the TIFF to minimize storage footprint.
// Use Cases: Archive large Excel reports as compact GZip‑packed TIFF files. | Attach size‑restricted TIFF images to emails after GZip compression. | Store massive numbers of rendered worksheets in cloud buckets with minimal storage cost.
// AI Prompts: Generate C# code that uses Aspose.Cells to render a worksheet to LZW‑compressed TIFF and then compress the file with GZip, including exception handling. | Explain how to switch ImageOrPrintOptions to other TIFF compression methods before applying GZip. | Show how to pipe the TIFF output directly into a GZipStream to avoid creating a temporary .tiff file.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, apply LZW compression via ImageOrPrintOptions, export the first worksheet as a TIFF file, and then shrink the TIFF further by compressing it with GZipStream, all with proper error handling.
class TiffGzipCompressionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample data for TIFF");
            worksheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure image options for TIFF rendering with LZW compression
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Specify TIFF compression type
                TiffCompression = TiffCompression.CompressionLZW,
                OnePagePerSheet = true
            };

            // Render the worksheet to a TIFF file
            string tiffFilePath = "output.tiff";
            SheetRender renderer = new SheetRender(worksheet, imgOptions);
            renderer.ToTiff(tiffFilePath);

            // Verify that the TIFF file was created before compression
            if (!File.Exists(tiffFilePath))
                throw new FileNotFoundException("Rendered TIFF file not found.", tiffFilePath);

            // Compress the generated TIFF file using GZip
            string gzFilePath = "output.tiff.gz";
            using (FileStream originalFile = new FileStream(tiffFilePath, FileMode.Open, FileAccess.Read))
            using (FileStream compressedFile = new FileStream(gzFilePath, FileMode.Create, FileAccess.Write))
            using (GZipStream gzipStream = new GZipStream(compressedFile, CompressionLevel.Optimal))
            {
                originalFile.CopyTo(gzipStream);
            }

            Console.WriteLine("TIFF file rendered and compressed to GZip successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
