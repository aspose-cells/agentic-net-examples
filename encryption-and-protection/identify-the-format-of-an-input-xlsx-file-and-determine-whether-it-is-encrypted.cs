using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DetectFileFormatAndEncryption
    {
        public static void Run(string filePath)
        {
            // Detect the file format and encryption status using Aspose.Cells utility
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

            // Output detected format
            Console.WriteLine($"Detected File Format Type: {info.FileFormatType}");

            // Output encryption status
            Console.WriteLine($"Is Encrypted: {info.IsEncrypted}");
        }

        // Example entry point
        public static void Main()
        {
            // Replace with the path to your XLSX (or any Excel) file
            string path = "sample.xlsx";

            Run(path);
        }
    }
}