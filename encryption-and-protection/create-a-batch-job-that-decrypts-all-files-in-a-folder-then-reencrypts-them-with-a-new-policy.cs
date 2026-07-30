// Title: C# Batch Decrypt & Re‑Encrypt Excel Workbooks with New Policy via Aspose.Cells
// Description: A C# utility that scans a directory for .xlsx files, uses Aspose.Cells.Metadata to open each workbook in decryption mode, writes a temporary unprotected copy, reloads it, applies a new password, StrongCryptographicProvider encryption and a 128‑bit key, then overwrites the original file. The script cleans up temporary files and logs any errors.
// Keywords: Aspose.Cells | C# batch decrypt Excel | re‑encrypt workbook | WorkbookMetadata | Excel encryption password rotation | StrongCryptographicProvider | 128‑bit key | metadata decryption | GitHub example | global
// Common Searches: how to batch decrypt Excel files with Aspose.Cells | change password for multiple .xlsx files programmatically | update encryption type for Excel workbooks .NET | Aspose.Cells decrypt and re‑encrypt folder of spreadsheets | rotate Excel encryption keys using C#
// Developer Intent: Create a C# batch job that decrypts every Excel workbook in a folder and re‑encrypts it with a new password and encryption settings using Aspose.Cells.
// Use Cases: Migrate legacy encrypted spreadsheets to a stronger algorithm across a document repository. | Rotate encryption passwords for compliance by processing all stored Excel files in one operation. | Automate re‑encryption of user‑uploaded spreadsheets before storing them in a secure cloud bucket.
// AI Prompts: Generate C# code that scans a folder for .xlsx files, decrypts each with WorkbookMetadata, and re‑encrypts using a specified password and StrongCryptographicProvider. | Provide best‑practice error handling and cleanup for a batch decryption/re‑encryption routine with Aspose.Cells. | Explain how to extend the script to support .xls and .xlsb formats and log results to a CSV file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

// A C# utility that scans a directory for .xlsx files, uses Aspose.Cells.Metadata to open each workbook in decryption mode, writes a temporary unprotected copy, reloads it, applies a new password, StrongCryptographicProvider encryption and a 128‑bit key, then overwrites the original file. The script cleans up temporary files and logs any errors.
class BatchDecryptEncrypt
{
    // Folder containing the Excel files to process
    private const string InputFolder = @"C:\ExcelFiles";

    // Temporary folder to store decrypted intermediate files
    private static readonly string TempFolder = Path.Combine(Path.GetTempPath(), "AsposeDecryptTemp");

    // New encryption password and options
    private const string NewPassword = "NewSecurePassword123";
    private const EncryptionType NewEncryptionType = EncryptionType.StrongCryptographicProvider;
    private const int NewKeyLength = 128; // 128, 256 or 40 bits

    static void Main()
    {
        // Ensure temporary folder exists
        Directory.CreateDirectory(TempFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(InputFolder, "*.xlsx"))
        {
            try
            {
                // ---------- Decrypt the file ----------
                // Load workbook metadata with Decryption flag
                MetadataOptions decryptOptions = new MetadataOptions(MetadataType.Decryption);
                WorkbookMetadata metadata = new WorkbookMetadata(filePath, decryptOptions);

                // Save the decrypted workbook to a temporary location
                string tempFile = Path.Combine(TempFolder, Path.GetFileName(filePath));
                metadata.Save(tempFile);

                // ---------- Re‑encrypt the file with new policy ----------
                // Load the decrypted workbook normally
                Workbook workbook = new Workbook(tempFile);

                // Apply new encryption settings
                workbook.Settings.Password = NewPassword;
                workbook.SetEncryptionOptions(NewEncryptionType, NewKeyLength);

                // Overwrite the original file with the newly encrypted version
                workbook.Save(filePath, SaveFormat.Xlsx);

                // Clean up temporary file
                File.Delete(tempFile);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }

        // Remove temporary folder if empty
        if (Directory.GetFiles(TempFolder).Length == 0)
        {
            Directory.Delete(TempFolder);
        }
    }
}
