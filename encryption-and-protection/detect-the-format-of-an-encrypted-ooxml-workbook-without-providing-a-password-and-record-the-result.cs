using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class DetectEncryptedWorkbookFormat
    {
        public static void Run()
        {
            try
            {
                // Path to the encrypted OOXML workbook (XLSX)
                string encryptedFilePath = "encrypted.xlsx";

                // Verify that the file exists before attempting detection
                if (!File.Exists(encryptedFilePath))
                {
                    Console.WriteLine($"File not found: {encryptedFilePath}");
                    return;
                }

                // Detect the file format without providing a password
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(encryptedFilePath);

                // Prepare a summary of detection results
                string result = $"File: {encryptedFilePath}{Environment.NewLine}" +
                                $"Detected Format: {info.FileFormatType}{Environment.NewLine}" +
                                $"Is Encrypted: {info.IsEncrypted}{Environment.NewLine}" +
                                $"Is Protected By RMS: {info.IsProtectedByRMS}{Environment.NewLine}" +
                                $"Load Format: {info.LoadFormat}";

                // Output the result to the console
                Console.WriteLine(result);

                // Record the result to a text file
                string outputPath = "DetectionResult.txt";
                File.WriteAllText(outputPath, result);
                Console.WriteLine($"Detection result saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DetectEncryptedWorkbookFormat.Run();
        }
    }
}