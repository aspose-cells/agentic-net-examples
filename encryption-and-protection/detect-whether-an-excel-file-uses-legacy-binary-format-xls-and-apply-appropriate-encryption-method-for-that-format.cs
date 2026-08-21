// Title: Detect legacy .xls format and apply encryption with Aspose.Cells for .NET
// Description: C# sample that uses Aspose.Cells to identify an Excel 97‑2003 (.xls) workbook, set a password, configure EncryptionType and key length, and save with XlsSaveOptions that encrypt document properties. For .xlsx/.xlsm files it applies standard password protection only.
// Keywords: Aspose.Cells detect xls format | legacy .xls encryption .NET | EncryptionType Aspose.Cells | XlsSaveOptions encrypt properties | Workbook password protection C# | FileFormatUtil DetectFileFormat | Excel 97-2003 encryption | Aspose.Cells encryption example
// Common Searches: how to check if an Excel file is .xls using Aspose.Cells | apply strong encryption to legacy Excel workbook Aspose.Cells | set encryption type and key length for .xls files | encrypt document properties when saving .xls with Aspose.Cells | C# detect Excel format and protect workbook
// Developer Intent: Identify whether a workbook is in the old .xls format and apply the correct encryption settings before saving.
// Use Cases: Secure a batch of legacy .xls files by assigning a password and specifying a 128‑bit StrongCryptographicProvider encryption. | Automatically protect mixed Excel collections: full encryption for .xls, simple password protection for .xlsx/.xlsm. | Log the detected format and encryption outcome for compliance auditing.
// AI Prompts: Generate C# code that detects .xls files with Aspose.Cells, applies EncryptionType.StrongCryptographicProvider and a 256‑bit key, and saves with encrypted document properties. | Write a script that scans a directory, encrypts each .xls file using a user‑supplied password and key length, and only password‑protects non‑xls files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // C# sample that uses Aspose.Cells to identify an Excel 97‑2003 (.xls) workbook, set a password, configure EncryptionType and key length, and save with XlsSaveOptions that encrypt document properties. For .xlsx/.xlsm files it applies standard password protection only.
    public class LegacyXlsEncryption
    {
        /// <param name="sourcePath">Path to the source Excel file.</param>
        /// <param name="destPath">Path where the encrypted file will be saved.</param>
        /// <param name="password">Password to protect the workbook.</param>
        /// <param name="encryptionType">Encryption algorithm to use (only relevant for .xls).</param>
        /// <param name="keyLength">Key length for the encryption algorithm (e.g., 128).</param>
        public static void ApplyEncryption(string sourcePath, string destPath, string password,
                                           EncryptionType encryptionType, int keyLength)
        {
            // Verify source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Detect the file format of the source workbook.
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            bool isLegacyXls = formatInfo.FileFormatType == FileFormatType.Excel97To2003;

            // Load the workbook (no password needed for loading because we are encrypting a fresh copy).
            Workbook workbook = new Workbook(sourcePath);

            // Set the workbook password (required for both .xls and .xlsx).
            workbook.Settings.Password = password;

            if (isLegacyXls)
            {
                // For legacy .xls files, specify encryption options.
                workbook.SetEncryptionOptions(encryptionType, keyLength);

                // Create XlsSaveOptions to control .xls specific saving behavior.
                XlsSaveOptions saveOptions = new XlsSaveOptions
                {
                    // Ensure document properties are also encrypted (default is true).
                    EncryptDocumentProperties = true
                };

                // Save the workbook as .xls with the specified options.
                workbook.Save(destPath, saveOptions);
            }
            else
            {
                // For modern formats (.xlsx, .xlsm, etc.), simple password protection is sufficient.
                // Save using the default format inferred from the destination file extension.
                workbook.Save(destPath);
            }

            // Output detection result.
            Console.WriteLine($"Source file format: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is legacy .xls: {isLegacyXls}");
            Console.WriteLine($"Encrypted file saved to: {destPath}");
        }

        // Example usage
        public static void Run()
        {
            try
            {
                string sourceFile = "input.xls";          // Replace with your source file path
                string encryptedFile = "output_encrypted.xls";
                string password = "MySecretPwd";

                // Choose an encryption type supported for Excel 97-2003 files.
                EncryptionType encType = EncryptionType.StrongCryptographicProvider;
                int keyLen = 128; // Valid values: 40, 128, 256

                ApplyEncryption(sourceFile, encryptedFile, password, encType, keyLen);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            LegacyXlsEncryption.Run();
        }
    }
}
