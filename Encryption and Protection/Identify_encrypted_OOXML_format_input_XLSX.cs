using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class DetectEncryptedOoxml
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to be examined.
            // You can replace this with any valid path or pass it as a command‑line argument.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Detect the file format and encryption status using the built‑in utility.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

            // Output the detected format.
            Console.WriteLine($"Detected File Format Type: {info.FileFormatType}");

            // Output whether the file is encrypted (requires a password to open).
            Console.WriteLine($"Is Encrypted: {info.IsEncrypted}");
        }
    }
}