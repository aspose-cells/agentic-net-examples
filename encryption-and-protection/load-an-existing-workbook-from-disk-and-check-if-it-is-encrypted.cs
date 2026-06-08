using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file to be examined
            string filePath = "sample.xlsx";

            // Verify that the file exists before attempting any operations
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return;
            }

            try
            {
                // Load the workbook without providing a password.
                // If the file is encrypted, Workbook.Settings.IsEncrypted will be true.
                Workbook workbook = new Workbook(filePath);
                bool isEncrypted = workbook.Settings.IsEncrypted;
                Console.WriteLine($"Workbook loaded from '{filePath}' IsEncrypted: {isEncrypted}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
            }

            try
            {
                // Detect file format (including encryption) without loading the whole workbook
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"FileFormatUtil reports IsEncrypted: {formatInfo.IsEncrypted}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to detect file format: {ex.Message}");
            }
        }
    }
}