using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook file (replace with your actual file path)
            string filePath = "example.xlsx";

            // Detect the file format and encryption status
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            bool isEncrypted = formatInfo.IsEncrypted;

            // Log encryption status
            Console.WriteLine($"Is the file encrypted? {isEncrypted}");

            // Load the workbook only if it is not encrypted
            if (!isEncrypted)
            {
                // Load the workbook using the file path constructor
                Workbook workbook = new Workbook(filePath);

                // Log successful load and detected format
                Console.WriteLine($"Workbook loaded successfully. Detected format: {formatInfo.FileFormatType}");
            }
            else
            {
                Console.WriteLine("The workbook is encrypted and cannot be loaded without a password.");
            }
        }
    }
}