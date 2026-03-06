using System;
using System.IO;
using Aspose.Cells;

namespace DetectEncryptedFileFormat
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the encrypted XLSX file
            string filePath = "encrypted.xlsx";

            // Password for the encrypted file (replace with actual password)
            string password = "test";

            // Ensure the file exists before attempting detection
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Open the file as a read‑only stream
            using (Stream stream = File.OpenRead(filePath))
            {
                try
                {
                    // Detect the file format using the provided password
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(stream, password);

                    // Output detection results
                    Console.WriteLine($"File Format Type : {info.FileFormatType}");
                    Console.WriteLine($"Is Encrypted     : {info.IsEncrypted}");
                    Console.WriteLine($"Is Password Valid: {FileFormatUtil.VerifyPassword(stream, password)}");
                }
                catch (Exception ex)
                {
                    // Handle any errors that occur during detection
                    Console.WriteLine($"Error detecting file format: {ex.Message}");
                }
            }
        }
    }
}