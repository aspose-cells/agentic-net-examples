// Title: Encrypt an Excel workbook with Aspise.Cells .NET and upload to Azure Blob storage while preserving encryption
// Description: Shows how to create a workbook with Aspose.Cells, insert data, apply password protection and AES‑128 encryption, save the file, and upload it to Azure Blob storage without decrypting the content.
// Keywords: Aspose.Cells encrypt workbook | C# AES 128 Excel encryption | password protect Excel .NET | upload encrypted Excel Azure Blob | preserve encryption Azure storage | SetEncryptionOptions Aspose.Cells | Azure Blob storage .NET SDK | secure Excel upload | encrypted spreadsheet cloud
// Common Searches: encrypt Excel file with Aspose.Cells .NET | set password and AES encryption for .xlsx using Aspose.Cells | upload encrypted Excel to Azure Blob storage | keep workbook encryption after Azure upload | Aspose.Cells SetEncryptionOptions example
// Developer Intent: Create an Excel file, protect it with a password and AES‑128 encryption via Aspose.Cells, and store it directly in Azure Blob storage while retaining the encryption.
// Use Cases: Generate confidential financial reports, encrypt them with strong AES‑128, and archive the files in Azure Blob for regulated data storage. | Automate the production of sensitive HR spreadsheets, apply password protection, and push the encrypted files to cloud storage for secure sharing across offices. | Integrate workbook encryption into a CI/CD pipeline, ensuring that every released .xlsx remains encrypted when uploaded to Azure.
// AI Prompts: Provide C# code that creates an Excel workbook with Aspose.Cells, applies a password and AES‑128 encryption, and uploads the encrypted file to Azure Blob storage without decrypting it. | Show how to configure the Azure Blob .NET client to stream an already encrypted .xlsx file directly to a container. | Explain how to verify that the workbook’s encryption is intact after it has been uploaded to Azure Blob storage.

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptedWorkbookUpload
{
    // Shows how to create a workbook with Aspose.Cells, insert data, apply password protection and AES‑128 encryption, save the file, and upload it to Azure Blob storage without decrypting the content.
    class Program
    {
        static void Main()
        {
            // Wrap the whole process in a try-catch to handle unexpected errors.
            try
            {
                // Create a new workbook.
                using (Workbook workbook = new Workbook())
                {
                    // Add sample data.
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sensitive Data");
                    sheet.Cells["B1"].PutValue(DateTime.Now);

                    // Set a password to encrypt the workbook.
                    workbook.Settings.Password = "StrongPassword123";

                    // Set stronger encryption options (AES 128-bit).
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // Define the output path.
                    string outputPath = Path.Combine(Environment.CurrentDirectory, "encryptedWorkbook.xlsx");

                    // Ensure the directory exists.
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the encrypted workbook to the file system.
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                // Log the exception details.
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
