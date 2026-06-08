using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFileFormatDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file whose format needs to be identified
            string filePath = "sample.xlsx";

            try
            {
                // Verify that the file exists before attempting to open it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
                    return;
                }

                // Open the file as a read‑only stream
                using (FileStream stream = File.OpenRead(filePath))
                {
                    // Detect the file format without loading the entire workbook
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);

                    // Output detected information
                    Console.WriteLine($"Detected LoadFormat   : {formatInfo.LoadFormat}");
                    Console.WriteLine($"Detected FileFormatType: {formatInfo.FileFormatType}");
                    Console.WriteLine($"Is Encrypted          : {formatInfo.IsEncrypted}");

                    // Create LoadOptions based on the detected LoadFormat (no workbook is loaded)
                    LoadOptions loadOptions = new LoadOptions(formatInfo.LoadFormat);
                    Console.WriteLine($"LoadOptions.LoadFormat: {loadOptions.LoadFormat}");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}