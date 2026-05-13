using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Prepare a source stream containing an Excel workbook.
        //    (In real scenarios the stream would come from a file,
        //     a database, a web request, etc.)
        // ------------------------------------------------------------
        MemoryStream sourceStream = new MemoryStream();
        Workbook sample = new Workbook();                     // create workbook (rule)
        sample.Worksheets[0].Cells["A1"].PutValue("Demo");   // add sample data
        sample.Save(sourceStream, SaveFormat.Xlsx);           // save to stream (rule)
        sourceStream.Position = 0;                           // reset for reading

        // ------------------------------------------------------------
        // 2. Load a workbook from the prepared stream.
        // ------------------------------------------------------------
        Workbook workbook = new Workbook(sourceStream);       // load from stream (rule)

        // ------------------------------------------------------------
        // 3. Configure OOXML save options with Level6 compression.
        // ------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); // create options (rule)
        saveOptions.CompressionType = OoxmlCompressionType.Level6; // set compression (rule)

        // ------------------------------------------------------------
        // 4. Save the workbook into a new MemoryStream using the options.
        // ------------------------------------------------------------
        using (MemoryStream resultStream = new MemoryStream())
        {
            workbook.Save(resultStream, saveOptions);        // save to stream with options (rule)

            // Reset position if the stream will be read later.
            resultStream.Position = 0;

            // Example output: display the size of the generated file.
            Console.WriteLine($"Workbook saved to MemoryStream, size = {resultStream.Length} bytes");
        }

        // ------------------------------------------------------------
        // 5. Clean up resources.
        // ------------------------------------------------------------
        sourceStream.Dispose();
        sample.Dispose();
        workbook.Dispose();
    }
}