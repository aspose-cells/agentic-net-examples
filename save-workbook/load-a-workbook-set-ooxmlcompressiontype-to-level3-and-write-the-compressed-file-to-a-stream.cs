// Title: Compress an Excel workbook with OOXML Level3 using Aspose.Cells for .NET and save to a MemoryStream
// Description: Load a workbook, set OoxmlSaveOptions.CompressionType to OoxmlCompressionType.Level3, and save the compressed Excel file into a MemoryStream. The stream can then be written to disk, sent in an HTTP response, or stored as a BLOB.
// Keywords: Aspose.Cells | OoxmlSaveOptions | OoxmlCompressionType.Level3 | compress Excel file | MemoryStream | C# | .NET | save workbook to stream | OOXML compression | Excel file size reduction
// Common Searches: Aspose.Cells set OOXML compression level | Save Excel workbook to MemoryStream with Level3 compression | How to compress an .xlsx file using Aspose.Cells for .NET | Write compressed workbook stream to file in C# | Available compression types in OoxmlSaveOptions
// Developer Intent: Load an existing workbook, apply OOXML Level3 compression, and output the compressed file to a stream for further processing.
// Use Cases: Generate a compressed Excel attachment in memory for email without creating a temporary file. | Stream a Level3‑compressed workbook directly to a web client for download. | Store a highly compressed workbook as a BLOB in a database to save storage space.
// AI Prompts: Show how to use OoxmlCompressionType.Maximum for the highest compression with Aspose.Cells. | Provide code to write a Level3‑compressed workbook directly to an HttpResponse stream. | Explain how to compare the original file size with the Level3 compressed stream programmatically.

using System;
using System.IO;
using Aspose.Cells;

// Load a workbook, set OoxmlSaveOptions.CompressionType to OoxmlCompressionType.Level3, and save the compressed Excel file into a MemoryStream. The stream can then be written to disk, sent in an HTTP response, or stored as a BLOB.
class Program
{
    static void Main()
    {
        // Load an existing workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create OOXML save options and set compression to Level3
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level3;

        // Save the workbook to a memory stream using the specified options
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, saveOptions);

            // Reset stream position if further processing is needed
            stream.Position = 0;

            // Example: write the compressed stream to a file
            using (FileStream file = new FileStream("compressed_output.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
        }
    }
}
