// Title: Compress HTML generated from an Excel workbook with Aspose.Cells using GZipStream in C#
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells, saves it as HTML in a MemoryStream, and returns the HTML compressed with GZipStream as a byte array. | Create code that reads the HTML string produced by Aspose.Cells, encodes it to UTF‑8 bytes, compresses those bytes using GZipStream, and outputs the compressed data for efficient web transmission.
// Common Searches: how to use Aspose.Cells to export Excel as HTML and gzip the output in .NET | C# compress HTML string from workbook with GZipStream for web API | Aspose.Cells save as HTML then compress bytes for HTTP response | generate gzip‑compressed HTML from Excel file using C# memory streams | best practice for sending Aspose.Cells HTML output as gzip in ASP.NET
// Tags: Aspose.Cells HTML export compression | GZipStream compress HTML bytes C# | Excel to HTML conversion Aspose.Cells | MemoryStream HTML generation .NET | Web delivery compressed HTML Aspose

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, saves it as HTML into a MemoryStream, reads the HTML as a UTF‑8 string, compresses the resulting byte array with GZipStream, and returns the compressed data for efficient web delivery.
    public static class HtmlCompressor
    {
        /// <param name="workbookPath">Full path to the source workbook.</param>
        /// <returns>Byte array containing the compressed HTML.</returns>
        public static byte[] GenerateCompressedHtml(string workbookPath)
        {
            try
            {
                // Verify that the workbook file exists.
                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException("Workbook file not found.", workbookPath);

                // Load the workbook using Aspose.Cells.
                var workbook = new Workbook(workbookPath);

                // Save the workbook as HTML into a memory stream.
                using (var htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, SaveFormat.Html);
                    htmlStream.Position = 0;

                    // Read the generated HTML as a UTF‑8 string.
                    string html;
                    using (var reader = new StreamReader(htmlStream, Encoding.UTF8))
                    {
                        html = reader.ReadToEnd();
                    }

                    // Convert HTML string to UTF‑8 bytes.
                    byte[] htmlBytes = Encoding.UTF8.GetBytes(html);

                    // Compress the HTML bytes using GZip.
                    using (var compressedStream = new MemoryStream())
                    {
                        using (var gzip = new GZipStream(compressedStream, CompressionMode.Compress, leaveOpen: true))
                        {
                            gzip.Write(htmlBytes, 0, htmlBytes.Length);
                        }

                        // Return the compressed data.
                        return compressedStream.ToArray();
                    }
                }
            }
            catch (Exception)
            {
                // Rethrow to allow caller to handle or log as needed.
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Specify the path to the Excel file you want to convert.
                string workbookPath = "sample.xlsx";

                // Generate compressed HTML.
                byte[] compressedHtml = HtmlCompressor.GenerateCompressedHtml(workbookPath);

                Console.WriteLine($"Compressed size: {compressedHtml.Length} bytes");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
