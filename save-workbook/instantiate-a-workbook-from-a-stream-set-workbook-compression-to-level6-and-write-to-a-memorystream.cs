// Title: Load an Excel workbook from a stream, apply Ooxml Level 6 compression, and save it to a MemoryStream using Aspose.Cells for .NET
// AI Prompts: Read an XLSX file from a MemoryStream, configure OoxmlSaveOptions.CompressionType to Level6, and write the compressed workbook into a new MemoryStream. | Demonstrate how to instantiate a Workbook with a stream constructor, set high‑compression Ooxml options, and output the result without touching the file system. | Provide C# code that takes an input Excel stream, applies Level6 Ooxml compression, and returns the compressed data as a MemoryStream using Aspose.Cells.
// Common Searches: how to set Ooxml compression level to 6 when saving an Excel workbook from a stream in C# | Aspose.Cells save workbook to MemoryStream with high compression | load Excel file from MemoryStream and compress output using OoxmlSaveOptions | C# example for stream‑to‑stream Excel compression with Aspose.Cells | configure OoxmlSaveOptions for Level6 compression without creating a temporary file
// Tags: OoxmlSaveOptions Level6 compression Aspose.Cells | instantiate Workbook from stream C# | save workbook to MemoryStream with compression | stream‑based Excel compression .NET | high‑compression XLSX output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    // The example creates a simple workbook, writes it to a MemoryStream, loads it back using the Workbook(stream) constructor, sets OoxmlSaveOptions.CompressionType to Level6, and saves the compressed workbook into another MemoryStream, demonstrating stream‑to‑stream Excel compression with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare a source Excel stream (could be any existing stream)
            // ------------------------------------------------------------
            // For demonstration, create a simple workbook and save it to a memory stream.
            Workbook sourceWorkbook = new Workbook();
            Worksheet srcSheet = sourceWorkbook.Worksheets[0];
            srcSheet.Cells["A1"].PutValue("Sample");
            srcSheet.Cells["B1"].PutValue(123);

            // Save the source workbook to a temporary memory stream in XLSX format.
            using (MemoryStream sourceStream = new MemoryStream())
            {
                sourceWorkbook.Save(sourceStream, SaveFormat.Xlsx);
                sourceStream.Position = 0; // Reset for reading.

                // ------------------------------------------------------------
                // 2. Load a workbook from the stream using the Stream constructor.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook(sourceStream);

                // ------------------------------------------------------------
                // 3. Configure OoxmlSaveOptions with Level6 compression.
                // ------------------------------------------------------------
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                saveOptions.CompressionType = OoxmlCompressionType.Level6;

                // ------------------------------------------------------------
                // 4. Save the workbook to a new MemoryStream using the options.
                // ------------------------------------------------------------
                using (MemoryStream resultStream = new MemoryStream())
                {
                    workbook.Save(resultStream, saveOptions);
                    // The resultStream now contains the workbook compressed with Level6.

                    // (Optional) Reset position if you need to read from the beginning.
                    resultStream.Position = 0;

                    // Example: write the stream to a file to verify the output.
                    using (FileStream file = new FileStream("CompressedOutput.xlsx", FileMode.Create, FileAccess.Write))
                    {
                        resultStream.CopyTo(file);
                    }

                    Console.WriteLine("Workbook saved to MemoryStream with Level6 compression.");
                }
            }
        }
    }
}
