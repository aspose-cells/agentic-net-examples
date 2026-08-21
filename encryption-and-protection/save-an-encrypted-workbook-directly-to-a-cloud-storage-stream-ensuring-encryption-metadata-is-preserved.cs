// Title: Encrypt and Save an Excel Workbook Directly to a Cloud Storage Stream with Aspose.Cells for .NET
// Description: Creates a Workbook, adds sample data, applies a password, selects StrongCryptographicProvider 128‑bit encryption, enables document‑property encryption via OoxmlSaveOptions, and writes the protected XLSX straight to a cloud storage stream (e.g., Azure Blob, Amazon S3, Google Cloud). No local file is created.
// Keywords: Aspose.Cells | C# | encrypt workbook | password protection | StrongCryptographicProvider | 128-bit encryption | OoxmlSaveOptions | encrypt document properties | cloud storage stream | Azure Blob Storage | Amazon S3 | Google Cloud Storage | save to stream | in‑memory Excel | Workbook.Save
// Common Searches: How to encrypt an Excel file with Aspose.Cells and save to a cloud stream | Aspose.Cells save encrypted XLSX to Azure Blob | Set encryption algorithm and key length in Aspose.Cells .NET | Encrypt document properties when saving workbook | Write encrypted Excel workbook directly to Amazon S3 stream | Upload password‑protected XLSX to Google Cloud Storage using Aspose.Cells
// Developer Intent: Create a password‑protected workbook with specific encryption settings and write the encrypted XLSX directly to a cloud storage stream without creating a local file.
// Use Cases: Upload a StrongCryptographicProvider‑encrypted XLSX to Azure Blob Storage via a Blob output stream. | Generate an in‑memory encrypted Excel file and pipe it to an Amazon S3 upload stream. | Store a password‑protected workbook in Google Cloud Storage by passing the GCS write stream to Workbook.Save. | Provide encrypted Excel reports to clients through a secure HTTP response stream.
// AI Prompts: Show C# code using Aspose.Cells to set a workbook password, apply StrongCryptographicProvider 128‑bit encryption, enable document‑property encryption, and save the file to an Azure Blob output stream. | Demonstrate how to obtain an Amazon S3 upload stream and use Workbook.Save with OoxmlSaveOptions to write an encrypted XLSX. | Explain the steps to encrypt a workbook and stream it directly to Google Cloud Storage without writing a temporary file. | Provide a generic method that accepts any cloud SDK stream and saves an encrypted workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsEncryptionToCloud
{
    // Creates a Workbook, adds sample data, applies a password, selects StrongCryptographicProvider 128‑bit encryption, enables document‑property encryption via OoxmlSaveOptions, and writes the protected XLSX straight to a cloud storage stream (e.g., Azure Blob, Amazon S3, Google Cloud). No local file is created.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted Data");
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Set a password to protect the workbook (encryption metadata)
            workbook.Settings.Password = "StrongPassword123";

            // Optional: specify encryption algorithm and key length (for Excel 97‑2003 compatibility)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Configure save options for OOXML format (XLSX) and ensure document properties are encrypted
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            saveOptions.EncryptDocumentProperties = true; // default, kept for clarity

            // Obtain a stream that writes directly to cloud storage.
            // Replace the implementation of GetCloudStorageStream with the actual cloud SDK stream.
            using (Stream cloudStream = GetCloudStorageStream())
            {
                // Save the encrypted workbook to the cloud stream using the specified options
                workbook.Save(cloudStream, saveOptions);

                // After saving, reset the stream position if further processing is needed
                cloudStream.Position = 0;
            }

            // Clean up
            workbook.Dispose();
        }

        // Placeholder method representing acquisition of a writable cloud storage stream.
        // In a real scenario, this could be a stream from Azure Blob Storage, Amazon S3, Google Cloud Storage, etc.
        private static Stream GetCloudStorageStream()
        {
            // For demonstration purposes, use a MemoryStream.
            // Replace this with the actual cloud storage stream creation logic.
            return new MemoryStream();
        }
    }
}
