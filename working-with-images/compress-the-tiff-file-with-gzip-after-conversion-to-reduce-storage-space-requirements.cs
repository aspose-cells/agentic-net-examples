// Title: Convert an Excel worksheet to a multi-page TIFF and compress it with GZip using Aspose.Cells for .NET (C#)
// Common Searches: how to export an Excel sheet to a TIFF file with LZW compression using Aspose.Cells in C# | C# Aspose.Cells example for converting workbook to TIFF and then gzipping the image | compress generated TIFF from Aspose.Cells with GZip to reduce storage size | Aspose.Cells render worksheet to multi-page TIFF and delete intermediate file in .NET | optimal way to GZip a TIFF file created from Excel using Aspose.Cells C#
// Tags: Aspose.Cells TIFF export with LZW compression | C# GZip compression of TIFF images | Excel worksheet to multi-page TIFF conversion | remove intermediate TIFF after GZip in .NET | reduce storage size for Excel-to-TIFF workflow

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffGzipCompression
{
    // The sample loads an Excel workbook, renders the first worksheet to a TIFF image using LZW compression at 300 dpi, compresses the TIFF with GZip at the optimal level, optionally deletes the intermediate TIFF file, and includes error handling and directory creation for a complete Excel‑to‑TIFF‑to‑GZip workflow in C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string excelPath = "input.xlsx";
            string tiffPath = "output.tiff";
            string gzPath   = "output.tiff.gz";

            try
            {
                // Verify source Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(excelPath);
                Worksheet worksheet = workbook.Worksheets[0]; // render first sheet

                // Set up TIFF rendering options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,                 // output format
                    TiffCompression = TiffCompression.CompressionLZW,
                    TiffColorDepth = ColorDepth.Format24bpp,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Render the worksheet to a TIFF file (single or multi‑page)
                SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
                sheetRender.ToTiff(tiffPath);
                Console.WriteLine($"Generated TIFF: {tiffPath}");

                // Compress the generated TIFF with GZip
                try
                {
                    CompressFile(tiffPath, gzPath);
                    Console.WriteLine($"Compressed to GZip: {gzPath}");
                }
                catch (Exception gzEx)
                {
                    Console.WriteLine($"Error during GZip compression: {gzEx.Message}");
                }

                // Optionally delete the intermediate TIFF to save space
                try
                {
                    File.Delete(tiffPath);
                    Console.WriteLine($"Deleted intermediate TIFF: {tiffPath}");
                }
                catch (Exception delEx)
                {
                    Console.WriteLine($"Warning: Could not delete intermediate TIFF '{tiffPath}'. {delEx.Message}");
                }

                Console.WriteLine("TIFF conversion and GZip compression completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="sourceFile">Full path of the file to compress.</param>
        /// <param name="destFile">Full path where the compressed file will be written.</param>
        private static void CompressFile(string sourceFile, string destFile)
        {
            // Ensure source exists
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException($"Source file '{sourceFile}' not found.");

            // Create destination directory if it does not exist
            string destDir = Path.GetDirectoryName(destFile);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Perform compression
            using (FileStream sourceStream = new FileStream(sourceFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (FileStream destStream = new FileStream(destFile, FileMode.Create, FileAccess.Write, FileShare.None))
            using (GZipStream compressionStream = new GZipStream(destStream, CompressionLevel.Optimal))
            {
                sourceStream.CopyTo(compressionStream);
            }
        }
    }
}
