// Title: C# – Load an Excel workbook from a stream, set OoxmlCompressionLevel6, and save to MemoryStream using Aspose.Cells
// Description: Demonstrates how to create a workbook, write it to a source MemoryStream, load it back with the Workbook(Stream) constructor, configure OoxmlSaveOptions to use Level 6 compression, and save the compressed workbook into a new MemoryStream ready for further processing.
// Keywords: Aspose.Cells C# load workbook from stream | OoxmlSaveOptions Level6 compression | save workbook to MemoryStream | compress XLSX in .NET | Aspose.Cells memory stream example | C# Excel compression without disk
// Common Searches: Aspose.Cells set compression level 6 when saving | How to save Excel to MemoryStream with compression in C# | Load workbook from byte array Aspose.Cells | Compress XLSX file in memory using Aspose.Cells | OoxmlSaveOptions CompressionType example
// Developer Intent: Load an existing workbook from a stream, apply Level 6 OOXML compression, and write the compressed file to another MemoryStream.
// Use Cases: Generate a lightweight Excel report for API responses without creating temporary files. | Store compressed XLSX blobs in a database to reduce storage costs. | Transmit Excel data over low‑bandwidth networks by compressing in memory.
// AI Prompts: Write C# code that reads an Excel file from a byte array, compresses it with OoxmlCompressionType.Level6 using Aspose.Cells, and returns the compressed byte array. | Explain the impact of OoxmlSaveOptions.CompressionType on XLSX file size and performance in Aspose.Cells. | Show the correct sequence for resetting stream positions when loading and saving workbooks with Aspose.Cells to avoid corrupted output.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, write it to a source MemoryStream, load it back with the Workbook(Stream) constructor, configure OoxmlSaveOptions to use Level 6 compression, and save the compressed workbook into a new MemoryStream ready for further processing.
class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample workbook and save it to a memory stream.
        //    This stream will be used as the source for loading.
        // ------------------------------------------------------------
        MemoryStream sourceStream = new MemoryStream();
        Workbook sampleWorkbook = new Workbook();                     // Workbook()
        sampleWorkbook.Worksheets[0].Cells["A1"].PutValue("Hello"); // add sample data
        sampleWorkbook.Save(sourceStream, SaveFormat.Xlsx);          // Save(Stream, SaveFormat)
        sourceStream.Position = 0; // reset for reading

        // ------------------------------------------------------------
        // 2. Load a workbook from the existing stream.
        //    Uses the Workbook(Stream) constructor.
        // ------------------------------------------------------------
        Workbook workbook = new Workbook(sourceStream); // Workbook(Stream)

        // ------------------------------------------------------------
        // 3. Configure OoxmlSaveOptions to use Level6 compression.
        // ------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();               // OoxmlSaveOptions()
        saveOptions.CompressionType = OoxmlCompressionType.Level6; // set compression

        // ------------------------------------------------------------
        // 4. Save the workbook into a new MemoryStream with the specified options.
        //    Uses Save(Stream, SaveOptions).
        // ------------------------------------------------------------
        using (MemoryStream resultStream = new MemoryStream())
        {
            workbook.Save(resultStream, saveOptions); // Save(Stream, SaveOptions)
            resultStream.Position = 0; // ready for further processing

            Console.WriteLine($"Compressed workbook size: {resultStream.Length} bytes");
        }

        // ------------------------------------------------------------
        // Cleanup
        // ------------------------------------------------------------
        sourceStream.Dispose();
        sampleWorkbook.Dispose();
        workbook.Dispose();
    }
}
