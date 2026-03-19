using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to be examined.
            // Replace with the actual file path or pass it as a command‑line argument.
            string filePath = args.Length > 0 ? args[0] : "example.xlsx";

            // Detect the file format and retrieve encryption information.
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Report whether the file is encrypted.
            Console.WriteLine($"Is file encrypted? {fileInfo.IsEncrypted}");
        }
    }
}