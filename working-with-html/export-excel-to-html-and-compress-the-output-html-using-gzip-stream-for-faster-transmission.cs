// Title: Compress Excel‑to‑HTML output with GZIP using Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook with Aspose.Cells, export it to full‑featured HTML via HtmlSaveOptions, and write the HTML into a GZipStream to generate a .gz file for faster transmission.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | HtmlSaveOptions | GZipStream | compress HTML | HTML .gz file | memory stream | web performance
// Common Searches: Aspose.Cells export Excel to HTML C# | C# GZIP compress HTML generated from Excel | Save Excel as compressed HTML .gz | HtmlSaveOptions all data export example | Stream GZIP HTML in ASP.NET Core
// Developer Intent: Produce a GZIP‑compressed HTML representation of an Excel workbook for bandwidth‑efficient delivery or storage.
// Use Cases: Web API that receives an XLSX file, converts it to full‑style HTML, compresses it with GZIP, and returns the .gz payload to clients. | Archiving spreadsheet previews as compressed HTML files to reduce storage while preserving formatting, images, and styles. | Serving pre‑compressed HTML reports from a web server to minimize bandwidth usage for end‑users.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to HTML with all styles and then compress the output using GZipStream. | Explain how to configure HtmlSaveOptions for full data export (styles, images, formulas) before applying GZIP compression. | Show how to write the GZIP‑compressed HTML directly to an ASP.NET Core response stream after converting an Excel file.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsHtmlGzipExample
{
    // Shows how to load an Excel workbook with Aspose.Cells, export it to full‑featured HTML via HtmlSaveOptions, and write the HTML into a GZipStream to generate a .gz file for faster transmission.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options (optional customizations)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Export all data (including styles, images, etc.)
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

            // Save the workbook as HTML into a memory stream using the provided Save(Stream, SaveOptions) rule
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);

                // Reset the stream position to the beginning before reading its content
                htmlStream.Position = 0;

                // Create the output file that will contain the GZIP-compressed HTML
                using (FileStream compressedFile = new FileStream("output.html.gz", FileMode.Create, FileAccess.Write))
                {
                    // Wrap the file stream with GZipStream for compression
                    using (GZipStream gzip = new GZipStream(compressedFile, CompressionMode.Compress))
                    {
                        // Copy the HTML data into the GZIP stream
                        htmlStream.CopyTo(gzip);
                    }
                }
            }

            Console.WriteLine("Excel workbook has been exported to compressed HTML (output.html.gz).");
        }
    }
}
