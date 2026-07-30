// Title: Get Excel Encryption Algorithm Name with Aspose.Cells (C#)
// Description: C# helper that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify whether an Excel workbook is encrypted and returns the encryption algorithm name (e.g., SHA‑AES for OOXML files) or a status message for unencrypted, legacy, or missing files.
// Keywords: Aspose.Cells encryption detection | C# Excel encryption algorithm | FileFormatUtil DetectFileFormat | Excel file IsEncrypted check | OOXML SHA‑AES encryption | legacy Excel encryption | GetEncryptionAlgorithmName | Excel security audit C#
// Common Searches: how to read encryption algorithm of an xlsx file using Aspose.Cells | C# detect if Excel workbook is encrypted | retrieve encryption type of Excel file Aspose | Aspose.Cells get encryption algorithm name | determine OOXML encryption algorithm programmatically
// Developer Intent: Return the name of the encryption algorithm applied to a specified Excel file.
// Use Cases: Validate that uploaded workbooks use standard OOXML encryption before processing. | Log encryption details of incoming Excel files for compliance reporting. | Display a clear message in a UI indicating encryption status and algorithm.
// AI Prompts: Write a C# method that uses Aspose.Cells to detect encryption and return the algorithm name, handling modern OOXML and legacy formats. | Create unit tests for GetEncryptionAlgorithmName covering encrypted .xlsx, unencrypted .xls, missing file, and exception scenarios. | Explain why Aspose.Cells can only expose the algorithm name for OOXML files and not for legacy binary Excel formats.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelEncryptionDemo
{
    // C# helper that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify whether an Excel workbook is encrypted and returns the encryption algorithm name (e.g., SHA‑AES for OOXML files) or a status message for unencrypted, legacy, or missing files.
    public static class ExcelEncryptionHelper
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>Encryption algorithm name or a status message.</returns>
        public static string GetEncryptionAlgorithmName(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return "Invalid path";

            if (!File.Exists(filePath))
                return "File not found";

            try
            {
                // Detect file format information without loading the workbook.
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);

                // If the file is not encrypted, indicate that.
                if (!info.IsEncrypted)
                    return "None";

                // Determine the encryption algorithm based on the detected file format.
                switch (info.FileFormatType)
                {
                    case FileFormatType.Xlsx:
                    case FileFormatType.Xlsb:
                    case FileFormatType.Xlsm:
                    case FileFormatType.Xltx:
                    case FileFormatType.Xltm:
                        return "SHA‑AES (standard OOXML encryption)";

                    // Legacy binary formats (e.g., .xls) – exact algorithm not exposed.
                    default:
                        return "Legacy Excel encryption (algorithm not identifiable)";
                }
            }
            catch (Exception ex)
            {
                // Return error information; in production you might log the exception instead.
                return $"Error: {ex.Message}";
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage
            string path = "encrypted.xlsx";

            string algorithm = ExcelEncryptionHelper.GetEncryptionAlgorithmName(path);
            Console.WriteLine($"Encryption algorithm: {algorithm}");
        }
    }
}
