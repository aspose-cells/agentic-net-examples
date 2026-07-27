// Title: Detect legacy XLS format and encrypt with StrongCryptographicProvider using Aspose.Cells for .NET
// Description: This example checks if a file is an Excel 97‑2003 binary (XLS), loads it (optionally with an existing password), assigns a new password, applies StrongCryptographicProvider 128‑bit encryption, encrypts document properties, and saves the protected workbook as an XLS file.
// Keywords: Aspose.Cells encrypt XLS | detect Excel 97-2003 format C# | StrongCryptographicProvider encryption | XLS password protection .NET | XlsSaveOptions encrypt properties
// Common Searches: how to encrypt an XLS file with Aspose.Cells | detect legacy Excel format before applying protection | set password for Excel 97-2003 workbook using Aspose | encrypt document properties when saving XLS | C# Aspose.Cells encryption for legacy Excel files
// Developer Intent: Identify a legacy XLS file and apply password‑based StrongCryptographicProvider encryption with Aspose.Cells.
// Use Cases: Secure confidential XLS reports before distribution. | Upgrade existing XLS password protection to a stronger algorithm. | Automate batch processing that validates format and encrypts only legacy workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to detect an Excel file's format and encrypt it with StrongCryptographicProvider when it is an XLS. | Refactor the LegacyXlsEncryptionDemo to handle both XLS and XLSX, selecting the appropriate encryption method for each. | Explain how to change the encryption key size or provider for legacy XLS files in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example checks if a file is an Excel 97‑2003 binary (XLS), loads it (optionally with an existing password), assigns a new password, applies StrongCryptographicProvider 128‑bit encryption, encrypts document properties, and saves the protected workbook as an XLS file.
    public class LegacyXlsEncryptionDemo
    {
        // Detects if the file is a legacy XLS and applies encryption appropriate for that format.
        public static void Run(string inputFilePath, string outputFilePath, string password)
        {
            try
            {
                // Verify input file exists.
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Detect the file format.
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(inputFilePath);
                Console.WriteLine($"Detected format: {fileInfo.FileFormatType}");
                Console.WriteLine($"Is encrypted already: {fileInfo.IsEncrypted}");

                // Proceed only if the file is a legacy Excel 97-2003 binary file (XLS).
                if (fileInfo.FileFormatType == FileFormatType.Excel97To2003)
                {
                    // Load the workbook (use LoadOptions with password if it is already encrypted).
                    LoadOptions loadOptions = new LoadOptions();
                    if (fileInfo.IsEncrypted)
                    {
                        loadOptions.Password = password;
                    }

                    Workbook workbook = new Workbook(inputFilePath, loadOptions);

                    // Set the password that will protect the workbook.
                    workbook.Settings.Password = password;

                    // Apply encryption options suitable for XLS files.
                    // Using StrongCryptographicProvider with a 128‑bit key as an example.
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Create save options for XLS format.
                    XlsSaveOptions saveOptions = new XlsSaveOptions
                    {
                        // Ensure document properties are also encrypted (default is true).
                        EncryptDocumentProperties = true
                    };

                    // Ensure output directory exists.
                    string outputDir = Path.GetDirectoryName(outputFilePath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the encrypted workbook as an XLS file.
                    workbook.Save(outputFilePath, saveOptions);

                    Console.WriteLine($"Encrypted XLS file saved to: {outputFilePath}");
                }
                else
                {
                    Console.WriteLine("The provided file is not a legacy XLS format. No encryption applied.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during encryption: {ex.Message}");
            }
        }
    }

    // Entry point for the console application.
    public class Program
    {
        // Usage: LegacyXlsEncryptionDemo <inputFilePath> <outputFilePath> <password>
        public static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: LegacyXlsEncryptionDemo <inputFilePath> <outputFilePath> <password>");
                return;
            }

            string inputFilePath = args[0];
            string outputFilePath = args[1];
            string password = args[2];

            LegacyXlsEncryptionDemo.Run(inputFilePath, outputFilePath, password);
        }
    }
}
