// Title: Export an Excel workbook to HTML and gzip the output using Aspose.Cells in C#
// Common Searches: C# export Excel to HTML with Aspose.Cells and compress with GZipStream | Save Aspose.Cells workbook as a gzipped HTML file in .NET | Generate .html.gz from Excel using Aspose.Cells memory stream | Compress HTML output from Aspose.Cells for faster web delivery | Aspose.Cells HTML export with gzip compression in C#
// Tags: Aspose.Cells HTML export memory stream | GZipStream HTML compression C# | save workbook as gzipped HTML .NET | Excel to compressed HTML Aspose.Cells | generate .html.gz file from Excel

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace ExcelToHtmlGzip
{
    // The example loads an Excel workbook with Aspose.Cells, saves it as HTML into a MemoryStream, compresses the HTML using a GZipStream, and writes the gzipped content to an .html.gz file for efficient transmission.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path to the compressed HTML output
            string gzipHtmlPath = "output.html.gz";

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // Save the workbook as HTML into a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                // Export to HTML format
                workbook.Save(htmlStream, SaveFormat.Html);
                htmlStream.Position = 0; // Reset stream position for reading

                // Compress the HTML content using GZIP
                using (MemoryStream compressedStream = new MemoryStream())
                {
                    // Create GZIP stream for compression
                    using (GZipStream gzipStream = new GZipStream(compressedStream, CompressionMode.Compress, true))
                    {
                        // Copy HTML data into the GZIP stream
                        htmlStream.CopyTo(gzipStream);
                    }

                    // Get the compressed bytes
                    byte[] compressedData = compressedStream.ToArray();

                    // Write the compressed data to the output file
                    File.WriteAllBytes(gzipHtmlPath, compressedData);
                }
            }

            Console.WriteLine("Excel file has been exported to compressed HTML successfully.");
        }
    }
}
