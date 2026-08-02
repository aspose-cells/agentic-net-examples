// Title: Detect Encrypted XLSX Workbook Format Without a Password – Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file type and encryption state of an OOXML workbook (XLSX) without providing a password. The script prints the detection details to the console and writes them to a text file for logging or audit purposes.
// Keywords: Aspose.Cells | .NET | C# | detect encrypted XLSX | FileFormatUtil | OOXML workbook detection | no password | file format identification | encryption status | audit log
// Common Searches: Aspose.Cells detect encrypted XLSX without password | FileFormatUtil DetectFileFormat example C# | how to check if an Excel file is password protected using Aspose | save workbook format detection result to a file | determine OOXML workbook encryption state programmatically
// Developer Intent: Find out whether an XLSX file is encrypted and what format it uses without supplying a password, then record the information.
// Use Cases: Pre‑process incoming Excel uploads in a web API and reject encrypted files before further handling. | Create an audit trail that logs workbook format and encryption flags for compliance reporting. | Route password‑protected spreadsheets to a secure decryption workflow based on detection results.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect if an XLSX file is encrypted without a password and output the file format. | Show how to extract FileFormatInfo properties (FileFormatType, IsEncrypted, IsProtectedByRMS, LoadFormat) and write them to a log file. | Explain error handling for missing files and unexpected exceptions when calling FileFormatUtil.DetectFileFormat.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file type and encryption state of an OOXML workbook (XLSX) without providing a password. The script prints the detection details to the console and writes them to a text file for logging or audit purposes.
    public class DetectEncryptedWorkbookFormat
    {
        public static void Run()
        {
            try
            {
                // Path to the encrypted OOXML workbook (XLSX)
                string filePath = "encrypted.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Input file not found: {filePath}");
                    return;
                }

                // Detect the file format without providing a password
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

                // Build a result string with the detection details
                string result = $"File: {Path.GetFileName(filePath)}{Environment.NewLine}" +
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
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DetectEncryptedWorkbookFormat.Run();
        }
    }
}
