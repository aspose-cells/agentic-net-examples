using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure a file path is provided as a command‑line argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsEncryptionChecker <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Detect the file format and retrieve encryption information
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output whether the file is encrypted
            Console.WriteLine($"Is file encrypted? {fileInfo.IsEncrypted}");

            // Provide a friendly message based on the encryption status
            if (fileInfo.IsEncrypted)
            {
                Console.WriteLine("The file requires a password to open.");
            }
            else
            {
                Console.WriteLine("The file is not encrypted.");
            }
        }
    }
}