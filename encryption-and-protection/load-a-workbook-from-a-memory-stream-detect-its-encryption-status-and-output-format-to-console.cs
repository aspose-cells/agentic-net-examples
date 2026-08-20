// Title: Detect Excel Workbook Encryption from a MemoryStream with Aspose.Cells for .NET
// Description: Creates an in‑memory workbook, saves it to a MemoryStream, then uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file format and whether the workbook is password‑protected, printing the results to the console.
// Keywords: Aspose.Cells encryption detection | FileFormatUtil DetectFileFormat | MemoryStream Excel .NET | check Excel password protection C# | detect workbook format type
// Common Searches: Aspose.Cells detect encrypted workbook from stream | FileFormatUtil DetectFileFormat example C# | how to know if Excel file is password protected without saving | read Excel from MemoryStream and check encryption
// Developer Intent: Determine if a workbook loaded via a MemoryStream is encrypted and retrieve its format type using Aspose.Cells.
// Use Cases: Validate user‑uploaded spreadsheets in a web API before processing. | Log format and encryption status for compliance audits. | Skip or prompt for a password when an incoming file is flagged as encrypted.
// AI Prompts: Generate C# code that reads an Excel byte array, detects password protection with Aspose.Cells, and returns the file format. | Explain how to handle a workbook flagged as encrypted by FileFormatUtil, including prompting for a password or aborting the operation.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDetection
{
    // Creates an in‑memory workbook, saves it to a MemoryStream, then uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file format and whether the workbook is password‑protected, printing the results to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Optionally add some data to the workbook
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Save the workbook into a memory stream
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);

                // Reset the stream position to the beginning before reading
                stream.Position = 0;

                // Detect file format and encryption status from the stream
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);

                // Output detection results to the console
                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
            }

            // Clean up the workbook instance
            workbook.Dispose();
        }
    }
}
