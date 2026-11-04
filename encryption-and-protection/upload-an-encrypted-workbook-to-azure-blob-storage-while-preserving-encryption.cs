// Title: How to upload a password‑protected Excel workbook to Azure Blob storage without removing encryption using Aspose.Cells for .NET
// Common Searches: C# upload an encrypted .xlsx file to Azure Blob while keeping the password using Aspose.Cells | preserve Excel workbook password when saving to Azure Blob storage with Aspose.Cells LoadOptions | save a password‑protected workbook to a MemoryStream before Azure Blob upload in .NET | re‑apply workbook password in Aspose.Cells before uploading to Azure Blob storage | load encrypted Excel file and stream it directly to Azure Blob using Aspose.Cells
// Tags: Aspose.Cells LoadOptions password protected XLSX | Workbook.Save to MemoryStream Aspose.Cells | Azure Blob storage encrypted Excel upload | C# reapply workbook password Aspose.Cells | stream copy encrypted workbook to Azure storage | preserve Excel encryption during cloud upload

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading a password‑protected XLSX workbook with Aspose.Cells LoadOptions, re‑applying the same password via Workbook.Settings, saving the workbook to a MemoryStream, and then writing the stream to a file—ready for direct upload to Azure Blob storage while maintaining the workbook's encryption.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the encrypted workbook on local disk
            string localEncryptedFilePath = @"C:\Temp\EncryptedWorkbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(localEncryptedFilePath))
            {
                Console.WriteLine($"Error: File not found at '{localEncryptedFilePath}'.");
                return;
            }

            // Password used to encrypt the workbook
            string workbookPassword = "MySecretPassword";

            // Load the encrypted workbook using the password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = workbookPassword
            };
            var workbook = new Workbook(localEncryptedFilePath, loadOptions);

            // Set the password for saving (re‑encrypt the workbook)
            workbook.Settings.Password = workbookPassword;

            // Save the workbook to a memory stream
            using var memoryStream = new MemoryStream();
            workbook.Save(memoryStream, SaveFormat.Xlsx);
            memoryStream.Position = 0; // Reset stream position before further processing

            // Example: Save the encrypted workbook to a local output file
            string outputPath = @"C:\Temp\EncryptedWorkbook_Uploaded.xlsx";
            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                memoryStream.CopyTo(fileStream);
            }

            Console.WriteLine($"Encrypted workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
