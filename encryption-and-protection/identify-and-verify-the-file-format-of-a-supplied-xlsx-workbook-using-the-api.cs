using System;
using Aspose.Cells;

namespace AsposeCellsFileFormatDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook whose format we want to identify.
            // Replace with the actual path or pass as a command‑line argument.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Detect the file format using Aspose.Cells utility method.
            // This uses the FileFormatUtil.DetectFileFormat(string) rule.
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output detected information.
            Console.WriteLine($"Detected File Format Type : {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted               : {formatInfo.IsEncrypted}");
            Console.WriteLine($"Load Format                : {formatInfo.LoadFormat}");

            // Verify that the workbook is an XLSX file.
            if (formatInfo.FileFormatType == FileFormatType.Xlsx)
            {
                Console.WriteLine("Verification passed: The file is an XLSX workbook.");
            }
            else
            {
                Console.WriteLine("Verification failed: The file is not an XLSX workbook.");
            }
        }
    }
}