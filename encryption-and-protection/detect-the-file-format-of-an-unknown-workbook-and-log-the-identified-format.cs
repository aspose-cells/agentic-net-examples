using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            try
            {
                DetectWorkbookFormat.Run(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    public class DetectWorkbookFormat
    {
        // Detects the format of a workbook file and logs the results.
        public static void Run(string filePath)
        {
            try
            {
                // Prevent FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Detect file format using Aspose.Cells utility.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                Console.WriteLine($"Processing file: {filePath}");
                Console.WriteLine($"Detected FileFormatType : {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted           : {formatInfo.IsEncrypted}");
                Console.WriteLine($"Detected LoadFormat    : {formatInfo.LoadFormat}");

                // Load the workbook to demonstrate the Workbook.FileFormat property.
                using (Workbook wb = new Workbook(filePath))
                {
                    Console.WriteLine($"Workbook.FileFormat property reports: {wb.FileFormat}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in DetectWorkbookFormat: {ex.Message}");
            }
        }
    }
}