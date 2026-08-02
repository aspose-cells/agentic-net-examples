// Title: Encrypt and Save an Aspose.Cells Workbook Directly to a Cloud Stream (C#)
// Description: Create a workbook, apply a password with Workbook.Settings.Password, and save the encrypted file straight to a MemoryStream. The stream can be uploaded to Azure Blob, AWS S3, or Google Cloud Storage without writing a local file, preserving all encryption metadata.
// Keywords: Aspose.Cells encrypt workbook C# | save encrypted Excel to stream | Workbook.Settings.Password | cloud storage upload Aspose.Cells | Azure Blob Aspose.Cells | AWS S3 encrypted Excel | Google Cloud Storage Aspose.Cells | memory stream Excel encryption
// Common Searches: Aspose.Cells save password protected workbook to Azure Blob | How to upload encrypted Excel to AWS S3 using C# | Save encrypted workbook to MemoryStream Aspose.Cells | Preserve encryption metadata when streaming Excel to cloud
// Developer Intent: Encrypt a workbook in C# and write it directly to a cloud storage stream, avoiding intermediate files while keeping encryption metadata intact.
// Use Cases: Generate a password‑protected financial report in a web API and stream it to Azure Blob storage. | Create an encrypted spreadsheet in a background job and upload it to an Amazon S3 bucket. | Produce a secure Excel file in a serverless function and send the stream to Google Cloud Storage.
// AI Prompts: Provide C# code that creates an Aspose.Cells workbook, sets a password, and uploads the encrypted stream to Azure Blob storage. | Show how to configure Aspose.Cells encryption for older Excel formats before saving to a MemoryStream. | Write a method that receives a Stream and uploads it to AWS S3 after saving an encrypted workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCloudExample
{
    // Create a workbook, apply a password with Workbook.Settings.Password, and save the encrypted file straight to a MemoryStream. The stream can be uploaded to Azure Blob, AWS S3, or Google Cloud Storage without writing a local file, preserving all encryption metadata.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted");
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Set a password to protect the workbook (encryption metadata will be stored)
            workbook.Settings.Password = "StrongPassword123";

            // Optional: set encryption options for older Excel formats (ignored for .xlsx)
            // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Assume we have a cloud storage stream (e.g., from Azure Blob, AWS S3, etc.)
            // Here we use a MemoryStream as a placeholder for the cloud stream.
            using (MemoryStream cloudStream = new MemoryStream())
            {
                // Save the encrypted workbook directly to the stream in XLSX format
                // This uses the provided Save(Stream, SaveFormat) rule.
                workbook.Save(cloudStream, SaveFormat.Xlsx);

                // At this point the stream contains the encrypted workbook with its metadata.
                // Reset the position if the cloud SDK requires reading from the beginning.
                cloudStream.Position = 0;

                // Example placeholder: upload the stream to cloud storage.
                // UploadToCloud(cloudStream); // Implement according to your cloud provider.
                Console.WriteLine($"Workbook saved to stream. Size: {cloudStream.Length} bytes");
            }

            // Clean up
            workbook.Dispose();
        }

        // Placeholder method for actual cloud upload implementation.
        // static void UploadToCloud(Stream stream) { /* Cloud SDK upload logic */ }
    }
}
