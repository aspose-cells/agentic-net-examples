// Title: Compress a generated HTML string using GZipStream in C# for efficient web transmission
// AI Prompts: Write a C# method that takes an HTML string and returns a byte array compressed with GZipStream. | Show how to create a simple HTML document, compress it with GZipStream, and output the original and compressed byte counts to the console. | Adapt the compression routine to stream the GZip‑compressed HTML directly to an ASP.NET Core HttpResponse.
// Common Searches: how to use GZipStream to compress HTML content in C# | C# example measuring original and gzipped HTML size | compress generated HTML string for web delivery with .NET | write gzip compressed HTML to HTTP response in ASP.NET Core | memory stream vs file stream for gzip compression of HTML in C#
// Tags: gzipstream html compression c# | memory stream gzip compression .net | web payload gzip compression asp.net core | compare original vs gzipped html size .net | c# compress html string for efficient transmission

using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace AsposeCellsExample
{
    // The program builds a simple HTML string, compresses it with GZipStream into a byte array, and prints both the original and compressed byte counts, illustrating how to reduce HTML payload size for web transmission.
    class Program
    {
        static void Main(string[] args)
        {
            // Sample HTML content
            string html = "<!DOCTYPE html><html><head><title>Sample</title></head><body><h1>Hello, World!</h1></body></html>";

            // Compress the HTML
            byte[] compressedData = CompressHtml(html);

            // For demonstration: write compressed size
            Console.WriteLine($"Original size: {Encoding.UTF8.GetByteCount(html)} bytes");
            Console.WriteLine($"Compressed size: {compressedData.Length} bytes");
        }

        /// <param name="html">HTML content to compress.</param>
        /// <returns>Byte array containing the compressed data.</returns>
        public static byte[] CompressHtml(string html)
        {
            // Convert HTML string to UTF-8 bytes
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);

            using (var outputStream = new MemoryStream())
            {
                // Create GZipStream for compression
                using (var gzipStream = new GZipStream(outputStream, CompressionMode.Compress, leaveOpen: true))
                {
                    // Write the HTML bytes into the GZipStream
                    gzipStream.Write(htmlBytes, 0, htmlBytes.Length);
                }

                // Return the compressed bytes
                return outputStream.ToArray();
            }
        }
    }
}
