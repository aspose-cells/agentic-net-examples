// Title: Compress Aspose.Cells HTML Output with GZipStream in C#
// Description: Shows how to export an Aspose.Cells workbook to HTML, apply HtmlSaveOptions (active worksheet only, gridlines), write the result to a MemoryStream, and compress the HTML using .NET GZipStream for fast web delivery or compact storage.
// Keywords: Aspose.Cells | C# | HTML export | GZipStream | gzip HTML | HtmlSaveOptions | workbook to HTML | .NET compression | web performance | spreadsheet report compression
// Common Searches: Aspose.Cells export to HTML C# | gzip HTML Aspose.Cells | C# compress HTML file | How to use GZipStream with Aspose.Cells | Save workbook as compressed HTML .NET | HTML compression for spreadsheet reports
// Developer Intent: Generate HTML from a spreadsheet with Aspose.Cells and then compress the output via GZipStream to reduce payload size for transmission or archival.
// Use Cases: Web API endpoint that returns a gzipped HTML representation of a spreadsheet. | Scheduled job that creates compressed HTML reports for long‑term storage. | ASP.NET Core middleware that streams gzipped HTML directly to the client. | Saving compressed HTML files on disk to minimize storage requirements.
// AI Prompts: Write a reusable C# method that accepts a Workbook and returns a gzipped HTML byte array using Aspose.Cells and GZipStream. | Show how to configure HtmlSaveOptions to embed CSS and images before compressing the HTML. | Explain how to set the Content‑Encoding header and stream gzipped HTML in an ASP.NET Core controller. | Provide error‑handling patterns for large workbooks when compressing HTML output.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsGzipExample
{
    // Shows how to export an Aspose.Cells workbook to HTML, apply HtmlSaveOptions (active worksheet only, gridlines), write the result to a MemoryStream, and compress the HTML using .NET GZipStream for fast web delivery or compact storage.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Configure HTML save options (optional customizations)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExportGridLines = true
            };

            // Save the workbook as HTML into a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0; // Reset for reading

                // Prepare a stream to hold the GZip-compressed data
                using (FileStream compressedFile = new FileStream("output.html.gz", FileMode.Create, FileAccess.Write))
                using (GZipStream gzip = new GZipStream(compressedFile, CompressionLevel.Optimal))
                {
                    // Copy the HTML bytes into the GZip stream
                    htmlStream.CopyTo(gzip);
                }
            }

            Console.WriteLine("HTML content saved and compressed to 'output.html.gz'.");
        }
    }
}
