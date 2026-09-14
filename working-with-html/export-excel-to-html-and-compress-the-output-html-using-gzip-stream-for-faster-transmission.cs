// Title: Compress Aspose.Cells HTML export with GZipStream in C# for faster web delivery
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, saves it as HTML into a MemoryStream, and writes the result to a .gz file using GZipStream. | Show how to pipe the HTML output from Aspose.Cells directly into a GZipStream to produce a compressed .html.gz file for efficient transmission.
// Common Searches: How can I export an Excel workbook to compressed HTML using Aspose.Cells and .NET GZipStream? | C# example for saving Aspose.Cells HTML output as a .gz file | Compress HTML generated from Excel with GZipStream for web performance | Aspose.Cells save as HTML then gzip in C#
// Tags: Aspose.Cells HTML export with GZipStream | C# compress HTML output .gz | save workbook as HTML memory stream | gzip compressed HTML transmission .NET | export Excel to compressed HTML using Aspose

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The sample loads an Excel file via Aspose.Cells, saves it as HTML into a MemoryStream, then compresses that HTML with a GZipStream and writes the compressed data to a .gz file, enabling faster delivery over the network.
class ExcelToCompressedHtml
{
    static void Main(string[] args)
    {
        // Input Excel file path
        string excelPath = @"C:\Input\sample.xlsx";

        // Output compressed HTML file path
        string compressedHtmlPath = @"C:\Output\sample.html.gz";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(excelPath);

        // Save the workbook as HTML into a memory stream
        using (MemoryStream htmlStream = new MemoryStream())
        {
            // Export to HTML format
            workbook.Save(htmlStream, SaveFormat.Html);

            // Reset stream position to the beginning before reading
            htmlStream.Position = 0;

            // Create the output file stream for the compressed HTML
            using (FileStream fileStream = new FileStream(compressedHtmlPath, FileMode.Create, FileAccess.Write))
            {
                // Wrap the file stream with GZipStream for compression
                using (GZipStream gzipStream = new GZipStream(fileStream, CompressionMode.Compress))
                {
                    // Copy the HTML data into the GZipStream
                    htmlStream.CopyTo(gzipStream);
                }
            }
        }

        Console.WriteLine("Excel file has been exported to compressed HTML successfully.");
    }
}
