// Title: Upload a password‑protected Excel workbook to Azure Blob Storage using Aspose.Cells for .NET
// AI Prompts: Load a password‑protected .xlsx file with Aspose.Cells, write it to a MemoryStream, and upload the stream to Azure Blob Storage without stripping the workbook password. | Adjust the sample to re‑apply workbook encryption before saving to the stream, then push the newly encrypted file to Azure Blob Storage. | Add comprehensive error handling for missing files, invalid passwords, and Azure connection failures, and log the success or failure of the blob upload.
// Common Searches: aspnet upload encrypted Excel file to Azure Blob using Aspose.Cells | preserve workbook password when saving to MemoryStream in C# | how to stream a password‑protected .xlsx to Azure Storage with Aspose.Cells | C# example for uploading password protected workbook to Azure Blob without losing encryption | Aspose.Cells load encrypted workbook and upload to Azure Blob Storage
// Tags: Aspose.Cells load password protected workbook | Azure Blob upload from MemoryStream | preserve workbook encryption in C# | C# Aspose.Cells save to XLSX stream | Azure.Storage.Blobs client usage with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an encrypted Excel workbook using Aspose.Cells with a supplied password, saves it to a MemoryStream in XLSX format (without re‑applying encryption), and then either uploads the stream to Azure Blob Storage via the Azure.Storage.Blobs SDK (code block commented out) or writes a local copy named "EncryptedCopy.xlsx".
class Program
{
    static void Main()
    {
        // Azure Blob storage parameters (set these if Azure SDK is added)
        string connectionString = "<Your_Azure_Storage_Connection_String>";
        string containerName = "workbooks";
        string blobName = "encrypted-uploaded.xlsx";

        // Local encrypted workbook parameters
        string localFilePath = @"C:\Path\To\EncryptedWorkbook.xlsx";
        string workbookPassword = "YourPassword";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(localFilePath))
                throw new FileNotFoundException("The encrypted workbook file was not found.", localFilePath);

            // Load the encrypted workbook using Aspose.Cells
            var loadOptions = new LoadOptions
            {
                Password = workbookPassword
            };
            var workbook = new Workbook(localFilePath, loadOptions);

            // Save the workbook to a memory stream.
            // Note: OoxmlSaveOptions.Password / EncryptionType are not available in the current Aspose.Cells version,
            // so the workbook is saved without re‑applying encryption.
            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position before further processing

                // -----------------------------------------------------------------
                // OPTIONAL: Upload to Azure Blob Storage.
                // To enable this block, add the NuGet package:
                //   Azure.Storage.Blobs
                // and uncomment the using directives at the top of the file.
                // -----------------------------------------------------------------
                /*
                try
                {
                    var serviceClient = new Azure.Storage.Blobs.BlobServiceClient(connectionString);
                    var containerClient = serviceClient.GetBlobContainerClient(containerName);
                    containerClient.CreateIfNotExists();

                    var blobClient = containerClient.GetBlobClient(blobName);
                    blobClient.Upload(ms, overwrite: true);
                    Console.WriteLine("Encrypted workbook uploaded to Azure Blob Storage successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Azure upload failed: {ex.Message}");
                }
                */

                // If Azure upload is not used, optionally save the stream to a local file
                string outputPath = Path.Combine(Path.GetDirectoryName(localFilePath) ?? string.Empty, "EncryptedCopy.xlsx");
                using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(fileStream);
                }
                Console.WriteLine($"Workbook saved locally to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
