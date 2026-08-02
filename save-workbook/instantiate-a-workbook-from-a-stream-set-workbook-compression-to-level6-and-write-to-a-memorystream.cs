using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a simple workbook and save it to a memory stream.
            // ------------------------------------------------------------
            Workbook initialWorkbook = new Workbook();                     // create workbook
            Worksheet sheet = initialWorkbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(123);

            using (MemoryStream sourceStream = new MemoryStream())
            {
                // Save the workbook into the stream in XLSX format.
                initialWorkbook.Save(sourceStream, SaveFormat.Xlsx);
                sourceStream.Position = 0; // reset for reading

                // ------------------------------------------------------------
                // 2. Load a workbook from the existing stream.
                // ------------------------------------------------------------
                Workbook loadedWorkbook = new Workbook(sourceStream); // load from stream

                // ------------------------------------------------------------
                // 3. Prepare OoxmlSaveOptions with Level6 compression.
                // ------------------------------------------------------------
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                saveOptions.CompressionType = OoxmlCompressionType.Level6; // set compression level

                // ------------------------------------------------------------
                // 4. Save the loaded workbook to a new MemoryStream using the options.
                // ------------------------------------------------------------
                using (MemoryStream resultStream = new MemoryStream())
                {
                    loadedWorkbook.Save(resultStream, saveOptions); // save with compression
                    // The resultStream now contains the compressed XLSX file.
                    // It can be used further, e.g., written to a file or returned from a method.
                    Console.WriteLine($"Compressed workbook size: {resultStream.Length} bytes");
                }
            }
        }
    }
}