using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFileFormatDetection
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a sample workbook (using the provided Workbook constructor)
            Workbook sampleWorkbook = new Workbook();

            // Add some data to make the workbook non‑empty
            sampleWorkbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Step 2: Save the workbook into a MemoryStream (using the provided Save method)
            using (MemoryStream stream = new MemoryStream())
            {
                sampleWorkbook.Save(stream, SaveFormat.Xlsx);

                // Reset the stream position before reading
                stream.Seek(0, SeekOrigin.Begin);

                // Step 3: Detect the file format from the stream (using FileFormatUtil.DetectFileFormat)
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);

                // Step 4: Log the detected information
                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
                Console.WriteLine($"Detected Load Format: {formatInfo.LoadFormat}");
            }

            // Optional: keep console window open when run outside an IDE
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}