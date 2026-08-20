// Title: C# – Retrieve Excel Encryption Algorithm Name Using Aspose.Cells
// Description: A static C# method that verifies a file, detects its format with Aspose.Cells' FileFormatUtil, checks encryption status, and returns the algorithm name – “SHA AES” for OOXML workbooks (.xlsx, .xlsb, .xlsm), the specific EncryptionType or “XOR” for legacy .xls files, or appropriate messages for unencrypted or unknown files. Supports optional passwords and provides clear error output.
// Keywords: Aspose.Cells | C# encryption algorithm detection | Excel file encryption type | GetEncryptionAlgorithmName | FileFormatUtil | .NET | XLSX encryption | XLS legacy encryption | SHA AES | XOR encryption | Workbook password detection
// Common Searches: How to get the encryption algorithm of an Excel file with Aspose.Cells | C# detect Excel workbook encryption type .NET | Get encryption algorithm name for .xlsx or .xls using Aspose | Identify if an Excel file is encrypted and which algorithm is used | Aspose.Cells retrieve encryption algorithm for password‑protected workbook
// Developer Intent: Determine the encryption algorithm applied to a specific Excel workbook.
// Use Cases: Validate uploaded Excel files by reporting their encryption algorithm before processing. | Log encryption algorithms of password‑protected workbooks for compliance and audit trails. | Inform users when a workbook uses an unsupported or unknown encryption method. | Automate routing decisions based on encryption strength (e.g., SHA‑AES vs. legacy XOR).
// AI Prompts: Generate a C# unit test suite for GetEncryptionAlgorithmName covering encrypted .xlsx, encrypted .xls with password, and plain files. | Explain how to extend GetEncryptionAlgorithmName to recognize new encryption schemes introduced in future Excel versions. | Create a PowerShell script that calls the C# method to batch‑process a folder of Excel files and output their encryption algorithms.

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // A static C# method that verifies a file, detects its format with Aspose.Cells' FileFormatUtil, checks encryption status, and returns the algorithm name – “SHA AES” for OOXML workbooks (.xlsx, .xlsb, .xlsm), the specific EncryptionType or “XOR” for legacy .xls files, or appropriate messages for unencrypted or unknown files. Supports optional passwords and provides clear error output.
    public static class ExcelEncryptionHelper
    {
        /// <param name="filePath">Path to the Excel file.</param>
        /// <param name="password">
        /// Optional password required to open an encrypted file.
        /// If the file is encrypted and the password is unknown, the method will still
        /// report the algorithm based on the file format.
        /// </param>
        /// <returns>Name of the encryption algorithm or an error description.</returns>
        public static string GetEncryptionAlgorithmName(string filePath, string password = null)
        {
            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                return "File not found";
            }

            try
            {
                // Detect file format and encryption status
                FileFormatInfo formatInfo = string.IsNullOrEmpty(password)
                    ? FileFormatUtil.DetectFileFormat(filePath)
                    : FileFormatUtil.DetectFileFormat(filePath, password);

                // Not encrypted -> no algorithm
                if (!formatInfo.IsEncrypted)
                    return "None";

                // Determine algorithm based on file format
                switch (formatInfo.FileFormatType)
                {
                    // OOXML formats use SHA‑AES (the same algorithm Excel uses)
                    case FileFormatType.Xlsx:
                    case FileFormatType.Xlsb:
                    case FileFormatType.Xlsm:
                        return "SHA AES";

                    // Legacy binary format – try to obtain the specific EncryptionType
                    default:
                        // Detect legacy XLS by name to avoid compile‑time dependency on the enum value
                        if (formatInfo.FileFormatType.ToString().Equals("Xls", StringComparison.OrdinalIgnoreCase))
                        {
                            if (string.IsNullOrEmpty(password))
                            {
                                // Without a password we cannot load the workbook to inspect settings
                                return "XOR (default for legacy XLS)";
                            }

                            // Load the workbook with the supplied password
                            var loadOptions = new LoadOptions { Password = password };
                            var workbook = new Workbook(filePath, loadOptions);

                            // Attempt to read the EncryptionType property via reflection
                            PropertyInfo encTypeProp = workbook.Settings.GetType()
                                .GetProperty("EncryptionType", BindingFlags.Public | BindingFlags.Instance);

                            if (encTypeProp != null)
                            {
                                object encTypeValue = encTypeProp.GetValue(workbook.Settings);
                                return encTypeValue?.ToString() ?? "Unknown";
                            }

                            // Fallback if the property is unavailable
                            return "XOR (default for legacy XLS)";
                        }

                        // Other formats – algorithm not exposed
                        return "Unknown";
                }
            }
            catch (Exception ex)
            {
                // Return the exception message for diagnostic purposes
                return $"Error: {ex.Message}";
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // Example usage – replace with your actual file path and password if needed
                string filePath = args.Length > 0 ? args[0] : "sample.xlsx";
                string password = args.Length > 1 ? args[1] : null;

                string algorithm = ExcelEncryptionHelper.GetEncryptionAlgorithmName(filePath, password);
                Console.WriteLine($"Encryption algorithm: {algorithm}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
