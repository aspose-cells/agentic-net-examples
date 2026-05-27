using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be examined
            string filePath = "sample.xls";

            try
            {
                // Verify that the file exists before attempting to read it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File not found – \"{filePath}\"");
                    return;
                }

                // Detect file format information without loading the workbook
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);

                // Output encryption status
                Console.WriteLine($"Is encrypted: {fileInfo.IsEncrypted}");

                // Infer the encryption algorithm based on the detected file format
                string algorithm = "Unknown";

                if (fileInfo.IsEncrypted)
                {
                    // Use the string representation of the format to avoid enum version issues
                    string format = fileInfo.FileFormatType.ToString();

                    if (format.Equals("Xls", StringComparison.OrdinalIgnoreCase))
                    {
                        // Excel 97‑2003 (.xls) uses legacy XOR/compatible encryption
                        algorithm = "XOR/Compatible (Excel 97‑2003)";
                    }
                    else if (format.Equals("Xlsx", StringComparison.OrdinalIgnoreCase) ||
                             format.Equals("Xlsm", StringComparison.OrdinalIgnoreCase) ||
                             format.Equals("Xlsb", StringComparison.OrdinalIgnoreCase))
                    {
                        // Excel 2007+ (.xlsx, .xlsm, .xlsb) uses AES encryption
                        algorithm = "AES (Office Open XML)";
                    }
                }

                Console.WriteLine($"Encryption algorithm (inferred): {algorithm}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}