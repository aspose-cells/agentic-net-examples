using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Use first argument as file path or fallback to "sample.xlsx"
                string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

                // Verify that the file exists before attempting to read it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Read the Excel file into a byte array
                byte[] excelBytes = File.ReadAllBytes(filePath);

                // Load the bytes into a memory stream for format detection
                using (MemoryStream memoryStream = new MemoryStream(excelBytes))
                {
                    memoryStream.Position = 0; // Ensure stream is at the beginning

                    // Detect file format information directly from the stream
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(memoryStream);

                    // Report whether the workbook is encrypted
                    Console.WriteLine($"Workbook encrypted: {formatInfo.IsEncrypted}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}