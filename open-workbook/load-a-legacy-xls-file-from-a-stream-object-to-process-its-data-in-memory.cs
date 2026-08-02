// Title: Load a Legacy XLS File from a Stream and Process It In‑Memory with Aspose.Cells for .NET
// Description: Demonstrates how to open a legacy .xls workbook using a FileStream, optionally detect its format with FileFormatUtil, reset the stream, create a Workbook, read cell values, and save the workbook back to a MemoryStream—all without writing to disk.
// Keywords: Aspose.Cells load xls from stream | detect excel format stream .NET | read cell value legacy xls | save workbook to memory stream | in‑memory Excel processing | C# Aspose.Cells stream handling
// Common Searches: open legacy xls from stream Aspose.Cells | detect excel file format before loading .NET | reset stream after format detection Aspose | save Aspose.Cells workbook to MemoryStream | process xls in memory C#
// Developer Intent: Load a legacy .xls workbook from any Stream, read or modify its data, and optionally write the result to a MemoryStream.
// Use Cases: Read an old .xls file received from a web service or database without creating a temporary file. | Validate the format of an incoming Excel stream before loading it with Aspose.Cells. | Extract specific cell values (e.g., A1) for business logic processing. | Modify worksheets and return the updated workbook as a byte array for API responses.
// AI Prompts: Write C# code that opens a legacy .xls file from a Stream using Aspose.Cells, detects the file format, resets the Stream, reads cell A1, and saves the workbook to a MemoryStream. | Explain why resetting the Stream position is required after calling FileFormatUtil.DetectFileFormat. | Show how to use Workbook.SaveToStream to obtain the XLS data as a byte array after making changes.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to open a legacy .xls workbook using a FileStream, optionally detect its format with FileFormatUtil, reset the stream, create a Workbook, read cell values, and save the workbook back to a MemoryStream—all without writing to disk.
    class LoadLegacyXlsFromStream
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "legacy.xls";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Open the legacy XLS file as a stream
            using (Stream inputStream = File.OpenRead(inputPath))
            {
                // OPTIONAL: Detect the file format from the stream
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputStream);
                Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");

                // Reset the stream position after detection
                inputStream.Seek(0, SeekOrigin.Begin);

                // Load the workbook from the stream (legacy XLS is supported)
                Workbook workbook = new Workbook(inputStream);

                // Example processing: read a value from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");

                // OPTIONAL: Save the workbook back to a memory stream using the provided rule
                using (MemoryStream outputStream = workbook.SaveToStream())
                {
                    // outputStream now contains the workbook data in XLS format
                    Console.WriteLine($"Workbook saved to memory stream. Length: {outputStream.Length} bytes.");
                }
            }
        }
    }
}
