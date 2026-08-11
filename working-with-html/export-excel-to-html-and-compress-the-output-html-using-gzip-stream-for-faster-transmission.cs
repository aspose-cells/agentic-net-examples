// Title: Export Excel to Gzipped HTML with Aspose.Cells in C#
// Description: This example demonstrates loading an Excel workbook with Aspose.Cells, saving it as HTML via HtmlSaveOptions into a MemoryStream, and compressing the HTML using .NET's GZipStream to create a .gz file. The technique reduces bandwidth and storage and can be streamed directly to an HTTP response.
// Keywords: Aspose.Cells | C# HTML export | GZipStream compression | Excel to HTML | gzipped HTML | HtmlSaveOptions | memory stream | ASP.NET Core response | reduce bandwidth | file compression .NET
// Common Searches: Aspose.Cells export Excel to HTML C# | Compress Aspose.Cells HTML output with GZIP | Save Excel as gzipped HTML .NET | How to stream gzipped HTML from Aspose.Cells | HtmlSaveOptions memory stream example | C# GZipStream for HTML files
// Developer Intent: Create a gzipped HTML file from an Excel workbook using Aspose.Cells.
// Use Cases: Serve compressed HTML previews of uploaded Excel files in a web API to speed up client downloads. | Archive Excel reports as small .gz HTML files for long‑term storage. | Integrate the compression step into an ASP.NET Core controller that writes the GZIP stream directly to the HTTP response.
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, converts it to HTML using HtmlSaveOptions, and writes the result to a .gz file via GZipStream. | Show how to return the gzipped HTML produced by Aspose.Cells as an ASP.NET Core FileResult. | Explain how to modify HtmlSaveOptions to export only the active worksheet before compressing the HTML output.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// This example demonstrates loading an Excel workbook with Aspose.Cells, saving it as HTML via HtmlSaveOptions into a MemoryStream, and compressing the HTML using .NET's GZipStream to create a .gz file. The technique reduces bandwidth and storage and can be streamed directly to an HTTP response.
class ExcelToCompressedHtml
{
    static void Main()
    {
        // Load the Excel workbook from a file (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options (create rule)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Example: export the whole workbook (default), you can customize other options here
        htmlOptions.ExportActiveWorksheetOnly = false;

        // Save the workbook as HTML into a memory stream (save rule)
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            // Reset the stream position to the beginning before reading
            htmlStream.Position = 0;

            // Compress the HTML content using GZIP and write to a .gz file
            using (FileStream compressedFile = new FileStream("output.html.gz", FileMode.Create, FileAccess.Write))
            using (GZipStream gzipStream = new GZipStream(compressedFile, CompressionMode.Compress))
            {
                // Copy the HTML stream into the GZIP stream
                htmlStream.CopyTo(gzipStream);
            }
        }

        Console.WriteLine("Excel file has been exported to compressed HTML (output.html.gz).");
    }
}
