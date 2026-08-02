using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure a file path is provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsEncryptionChecker <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Verify that the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found - {filePath}");
                return;
            }

            try
            {
                // Detect file format and encryption status
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

                Console.WriteLine($"File: {filePath}");
                Console.WriteLine($"Is encrypted: {fileInfo.IsEncrypted}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the file: {ex.Message}");
            }
        }
    }
}