using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates detection of workbook file format and encryption status.
    public class Program
    {
        public static void Main(string[] args)
        {
            // Expect a file path as the first argument.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an Excel file as an argument.");
                return;
            }

            string filePath = args[0];
            DetectWorkbookInfo(filePath);
        }

        // Detects and prints format information for the specified file.
        private static void DetectWorkbookInfo(string filePath)
        {
            // Use Aspose.Cells utility to detect file format.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

            // Output detected information.
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Detected Format      : {info.FileFormatType}");
            Console.WriteLine($"Is Encrypted         : {info.IsEncrypted}");
            Console.WriteLine($"Is Protected By RMS  : {info.IsProtectedByRMS}");
            Console.WriteLine($"Load Format          : {info.LoadFormat}");
        }
    }
}