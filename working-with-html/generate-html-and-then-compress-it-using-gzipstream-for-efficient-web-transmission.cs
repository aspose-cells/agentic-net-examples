// Title: Compress Aspose.Cells HTML Export with GZipStream in C# – Fast Web Delivery
// Description: A complete C# example that creates an Aspose.Cells workbook, exports the active worksheet to HTML using HtmlSaveOptions, compresses the HTML with GZipStream via MemoryStream, and writes the result to a .gz file or streams it in an ASP.NET response for efficient bandwidth usage.
// Keywords: Aspose.Cells HTML export C# | GZipStream compression .NET | C# memory stream gzip | export worksheet to HTML | gzip HTML response ASP.NET | compressed Excel report | Aspose.Cells HtmlSaveOptions | download gzipped HTML
// Common Searches: how to gzip Aspose.Cells HTML output in C# | C# code to compress Excel HTML export with GZipStream | Aspose.Cells export to HTML and send as gzip in ASP.NET Core | compress workbook HTML for web transmission .NET | save Aspose.Cells HTML as .gz file
// Developer Intent: Generate HTML from an Excel workbook and apply GZip compression for low‑latency delivery.
// Use Cases: Create a single‑sheet HTML report, gzip it, and return the byte array from a Web API endpoint. | Archive Excel‑derived HTML reports on disk as .gz files to save storage space. | Stream GZip‑compressed HTML directly to browsers with the appropriate Content‑Encoding header.
// AI Prompts: Write a C# method that takes an Aspose.Cells Workbook and returns a GZip‑compressed HTML byte array. | Show how to configure HtmlSaveOptions to minimize HTML size before applying GZipStream. | Provide an ASP.NET Core controller action that streams the compressed HTML from a MemoryStream with the correct Content‑Encoding header.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsHtmlGzipDemo
{
    // A complete C# example that creates an Aspose.Cells workbook, exports the active worksheet to HTML using HtmlSaveOptions, compresses the HTML with GZipStream via MemoryStream, and writes the result to a .gz file or streams it in an ASP.NET response for efficient bandwidth usage.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells HTML Export");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue("Compressed HTML");

            // Configure HTML save options (using the provided constructor)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportActiveWorksheetOnly = true; // export only the first sheet
            htmlOptions.ExcludeUnusedStyles = true;       // reduce HTML size

            // Save the workbook as HTML into a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0; // reset for reading

                // Prepare a stream to hold the GZip-compressed data
                using (MemoryStream compressedStream = new MemoryStream())
                {
                    // Compress the HTML stream using GZipStream
                    using (GZipStream gzip = new GZipStream(compressedStream, CompressionMode.Compress, leaveOpen: true))
                    {
                        htmlStream.CopyTo(gzip);
                    }

                    // After compression, reset position to read the compressed bytes
                    compressedStream.Position = 0;

                    // Optionally write the compressed data to a .gz file for verification
                    using (FileStream file = new FileStream("output.html.gz", FileMode.Create, FileAccess.Write))
                    {
                        compressedStream.CopyTo(file);
                    }

                    Console.WriteLine("HTML content has been compressed and saved as 'output.html.gz'.");
                }
            }

            // Clean up
            workbook.Dispose();
        }
    }
}
