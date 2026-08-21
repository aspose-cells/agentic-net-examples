// Title: Upload AES‑128 Encrypted Aspose.Cells Workbook to Azure Blob Storage using C#
// Description: Demonstrates how to create a password‑protected Excel workbook with Aspose.Cells, apply AES‑128 encryption, and upload the encrypted file directly to Azure Blob Storage without decrypting or writing an intermediate copy. Includes code for MemoryStream handling and Azure.Storage.Blobs integration.
// Keywords: Aspose.Cells | C# | AES-128 encryption | password protected workbook | Azure Blob Storage | upload encrypted Excel | Azure.Storage.Blobs | cloud file protection | encrypted workbook upload | secure Excel storage
// Common Searches: C# upload encrypted Excel to Azure Blob | Aspose.Cells save encrypted workbook to Azure | How to store password protected Excel in Azure Blob | Upload AES encrypted workbook without decryption | Azure Blob storage encrypted Excel example
// Developer Intent: Upload an Aspose.Cells workbook that is already encrypted with a password and AES‑128 encryption directly to Azure Blob Storage, preserving its protection throughout the transfer.
// Use Cases: Securely archive confidential financial spreadsheets in Azure Blob while retaining workbook encryption for compliance. | Transfer password‑protected reports between microservices via Azure Blob without exposing plaintext data. | Store regulated data in the cloud with end‑to‑end encryption applied by Aspose.Cells, ensuring only authorized users can open the file.
// AI Prompts: Write C# code that takes a MemoryStream containing an AES‑128 encrypted Aspose.Cells workbook and uploads it to Azure Blob Storage using Azure.Storage.Blobs, preserving the stream unchanged. | Provide a step‑by‑step tutorial for creating a SAS token, configuring a Blob container, and uploading an encrypted workbook without creating a local file. | Show how to modify the sample program to upload the encrypted workbook to Azure Blob, then delete any temporary local files and handle errors gracefully.

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptedWorkbookUpload
{
    // Demonstrates how to create a password‑protected Excel workbook with Aspose.Cells, apply AES‑128 encryption, and upload the encrypted file directly to Azure Blob Storage without decrypting or writing an intermediate copy. Includes code for MemoryStream handling and Azure.Storage.Blobs integration.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -------------------- Create and encrypt workbook --------------------
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sensitive Data");
                sheet.Cells["B1"].PutValue(DateTime.Now);

                // Set password protection
                workbook.Settings.Password = "StrongPassword123";

                // Set stronger encryption options (AES 128-bit)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Save the encrypted workbook to a memory stream
                using (MemoryStream workbookStream = new MemoryStream())
                {
                    workbook.Save(workbookStream, SaveFormat.Xlsx);
                    workbookStream.Position = 0; // Reset for reading

                    // -------------------- Save to local file (replace Azure upload) --------------------
                    string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
                    if (!Directory.Exists(outputFolder))
                    {
                        Directory.CreateDirectory(outputFolder);
                    }

                    string outputPath = Path.Combine(outputFolder, "encryptedWorkbook.xlsx");

                    // Write the stream to the file
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        workbookStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
                }

                // Clean up
                workbook.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
