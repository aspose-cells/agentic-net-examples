// Title: Detect and Decrypt Encrypted OOXML (XLSX) Files with Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells to identify an encrypted .xlsx, verify the password, load the workbook via LoadOptions, and save a plain‑format copy. Guarantees compatibility with standard OOXML after decryption.
// Keywords: Aspose.Cells detect encrypted Excel | FileFormatUtil DetectFileFormat password | LoadOptions password decryption | remove Excel encryption C# | save unencrypted XLSX Aspose | verify Excel password .NET
// Common Searches: How to check if an Excel file is password protected using Aspose.Cells | C# code to decrypt a protected .xlsx with Aspose.Cells | Verify password of encrypted workbook Aspose.Cells .NET | Load encrypted Excel without knowing format Aspose.Cells | Batch remove encryption from XLSX files in C#
// Developer Intent: Determine the file type and encryption state of an OOXML workbook, confirm the password, decrypt it, and produce an unprotected .xlsx.
// Use Cases: Pre‑process incoming Excel uploads and skip encrypted files that require a password. | Validate user‑supplied passwords before opening protected workbooks. | Automate conversion of password‑protected spreadsheets to standard format for downstream analytics.
// AI Prompts: Write C# code that uses Aspose.Cells to detect whether a .xlsx is encrypted and return its encryption status. | Create a method that accepts an encrypted Excel path and password, verifies the password, loads the workbook, and saves it without encryption. | Explain best‑practice exception handling when decrypting an Excel file with Aspose.Cells in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that uses Aspose.Cells to identify an encrypted .xlsx, verify the password, load the workbook via LoadOptions, and save a plain‑format copy. Guarantees compatibility with standard OOXML after decryption.
    public class DetectAndDecryptEncryptedOoxml
    {
        public static void Run()
        {
            // Path to the encrypted OOXML file (e.g., .xlsx)
            string encryptedFilePath = "encrypted.xlsx";

            // Verify that the input file exists
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"Error: File not found - {encryptedFilePath}");
                return;
            }

            // Password used to protect the file
            string password = "test";

            try
            {
                // -------------------------------------------------
                // 1. Detect the file format and encryption status
                // -------------------------------------------------
                // DetectFileFormat overload that accepts a password is used for encrypted files
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedFilePath, password);

                Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

                // Optional: Verify that the supplied password is correct
                using (Stream verifyStream = File.OpenRead(encryptedFilePath))
                {
                    bool passwordValid = FileFormatUtil.VerifyPassword(verifyStream, password);
                    Console.WriteLine($"Password valid: {passwordValid}");
                }

                // -------------------------------------------------
                // 2. Load the workbook using the correct password
                // -------------------------------------------------
                // LoadOptions allows us to specify the password for decryption
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };

                // Load the workbook; Aspose.Cells will decrypt it internally
                using (Workbook workbook = new Workbook(encryptedFilePath, loadOptions))
                {
                    // Verify that the workbook reports being encrypted (should be true before decryption)
                    Console.WriteLine($"Workbook Settings.IsEncrypted (after load): {workbook.Settings.IsEncrypted}");

                    // -------------------------------------------------
                    // 3. Save the workbook to a new file to ensure compatibility
                    // -------------------------------------------------
                    // Saving without a password removes encryption, producing a standard OOXML file
                    string decryptedFilePath = "decrypted_copy.xlsx";
                    workbook.Save(decryptedFilePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Decrypted workbook saved to: {decryptedFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DetectAndDecryptEncryptedOoxml.Run();
        }
    }
}
