// Title: Get Excel File Encryption Algorithm with Aspose.Cells (.NET)
// Description: C# helper that validates a file path, uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine if the workbook is encrypted, and returns the algorithm name without loading the workbook. Returns "SHA‑AES" for OOXML files (.xlsx, .xlsb, .xlsm), a generic legacy description for .xls, "None" for unencrypted files, and throws clear exceptions for invalid input.
// Keywords: Aspose.Cells encryption detection | C# Excel encryption algorithm | FileFormatUtil DetectFileFormat | Excel file encryption type .NET | Get encryption algorithm name | Excel security Aspose | OOXML SHA‑AES | legacy XLS encryption | detect encrypted workbook
// Common Searches: asp.net get encryption algorithm of an xlsx file | detect if Excel workbook is encrypted using Aspose.Cells | c# retrieve encryption type of .xls file | how to know which algorithm protects an Excel file | asp.net check encryption algorithm without opening workbook
// Developer Intent: Retrieve the encryption algorithm applied to a specific Excel file.
// Use Cases: Validate that uploaded workbooks use only approved encryption methods before processing. | Log encryption algorithms across a batch of files for compliance audits. | Display the encryption algorithm in a UI when a protected workbook is opened.
// AI Prompts: Write a C# method with Aspose.Cells that returns the encryption algorithm name for a given Excel file path, handling missing files and errors. | Create unit tests for GetEncryptionAlgorithmName covering encrypted .xlsx, encrypted .xls, unencrypted files, and invalid paths. | Explain how FileFormatUtil.DetectFileFormat identifies encryption status and why the algorithm name is inferred from the file extension.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelEncryptionDemo
{
    // C# helper that validates a file path, uses Aspose.Cells FileFormatUtil.DetectFileFormat to determine if the workbook is encrypted, and returns the algorithm name without loading the workbook. Returns "SHA‑AES" for OOXML files (.xlsx, .xlsb, .xlsm), a generic legacy description for .xls, "None" for unencrypted files, and throws clear exceptions for invalid input.
    public static class ExcelEncryptionHelper
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>Encryption algorithm name or "None".</returns>
        public static string GetEncryptionAlgorithmName(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found.", filePath);

            try
            {
                // Detect file format and encryption status without loading the workbook.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                // If the file is not encrypted, simply return "None".
                if (!formatInfo.IsEncrypted)
                    return "None";

                // Determine the algorithm based on the file extension / format type.
                string extension = Path.GetExtension(filePath).ToLowerInvariant();

                switch (extension)
                {
                    case ".xlsx":
                    case ".xlsb":
                    case ".xlsm":
                        return "SHA‑AES (default for OOXML formats)";

                    case ".xls":
                        // Legacy format – could be XOR, Compatible, EnhancedCryptographicProviderV1,
                        // or StrongCryptographicProvider. The specific type is not exposed.
                        return "Legacy encryption (XOR/Compatible/EnhancedCryptographicProviderV1/StrongCryptographicProvider)";

                    default:
                        // Unknown extension but encrypted – return a generic message.
                        return "Encrypted (algorithm unknown)";
                }
            }
            catch (Exception ex)
            {
                // Wrap any exception with additional context.
                throw new InvalidOperationException($"Failed to determine encryption algorithm for file '{filePath}'.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Use first argument as file path or fallback to a default name.
            string filePath = args.Length > 0 ? args[0] : "encrypted.xlsx";

            try
            {
                string algorithm = ExcelEncryptionHelper.GetEncryptionAlgorithmName(filePath);
                Console.WriteLine($"Encryption algorithm: {algorithm}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
