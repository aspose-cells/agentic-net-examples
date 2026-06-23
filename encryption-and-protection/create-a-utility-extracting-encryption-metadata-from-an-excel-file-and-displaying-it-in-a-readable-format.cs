using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionMetadataUtility
{
    public static class EncryptionMetadataExtractor
    {
        /// <summary>
        /// Detects and displays encryption related metadata of the specified Excel file.
        /// </summary>
        /// <param name="filePath">Full path to the Excel file.</param>
        public static void DisplayEncryptionMetadata(string filePath)
        {
            // Detect file format and retrieve metadata information
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output readable information
            Console.WriteLine("=== Encryption Metadata ===");
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"File Format Type : {formatInfo.FileFormatType}");
            Console.WriteLine($"Load Format      : {formatInfo.LoadFormat}");
            Console.WriteLine($"Is Encrypted     : {formatInfo.IsEncrypted}");
            Console.WriteLine($"Is Protected by RMS : {formatInfo.IsProtectedByRMS}");
            Console.WriteLine("===========================");
        }

        // Example usage
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an Excel file as an argument.");
                return;
            }

            string excelPath = args[0];
            DisplayEncryptionMetadata(excelPath);
        }
    }
}