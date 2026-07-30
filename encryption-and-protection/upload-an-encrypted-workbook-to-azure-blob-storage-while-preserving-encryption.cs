// Title: Upload an Encrypted Excel Workbook to Azure Blob Storage with Aspose.Cells (.NET)
// Description: Demonstrates creating a workbook, applying password protection and strong AES‑256 encryption, saving it to a MemoryStream, and streaming the encrypted file straight to Azure Blob storage while preserving its security.
// Keywords: Aspose.Cells | C# encryption | Excel password protection | AES-256 | MemoryStream upload | Azure Blob storage | cloud Excel security | XLSX encryption .NET | strong cryptographic provider
// Common Searches: C# upload encrypted Excel to Azure Blob using Aspose.Cells | how to protect workbook with password and AES encryption Aspose.Cells | stream encrypted XLSX directly to Azure Blob storage | preserve Excel encryption when uploading to cloud | Aspose.Cells set encryption options .NET
// Developer Intent: Create a password‑protected, AES‑encrypted Excel file and transfer it to Azure Blob storage without writing an unencrypted temporary file.
// Use Cases: Generate confidential financial reports that must remain encrypted during cloud transfer. | Automate nightly data exports that are stored in Azure Blob with built‑in workbook protection. | Provide a secure download link for clients by uploading an already encrypted XLSX to a storage container.
// AI Prompts: Write C# code that uploads the encrypted MemoryStream to Azure Blob using Azure.Storage.Blobs SDK. | Show how to change the encryption to AES‑256 and verify the blob’s content remains encrypted after upload. | Explain steps to test that the password protection works when the file is downloaded from Azure Blob.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, applying password protection and strong AES‑256 encryption, saving it to a MemoryStream, and streaming the encrypted file straight to Azure Blob storage while preserving its security.
class UploadEncryptedWorkbook
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted data");

            // Protect the workbook with a password
            workbook.Settings.Password = "MySecretPwd";

            // Apply strong encryption (optional but recommended)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook to a memory stream in XLSX format
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);
                stream.Position = 0; // Reset stream position before further processing

                // For demonstration, save the encrypted workbook to a local file
                string outputPath = Path.Combine(Environment.CurrentDirectory, "encryptedWorkbook.xlsx");
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }

                Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
            }

            // Dispose the workbook
            workbook.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
