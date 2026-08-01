// Title: Validate Exported Certificate Size Against Original PFX in Aspose.Cells (C#)
// Description: Loads a .pfx certificate, records its byte length, signs an Aspose.Cells workbook, saves it to a MemoryStream, reloads the file, extracts the embedded certificate, and checks that the exported certificate size matches the original file size.
// Keywords: Aspose.Cells digital signature | C# certificate size validation | exported certificate length | compare PFX byte size | Excel workbook signing | Aspose.Cells certificate export | certificate byte count check
// Common Searches: Aspose.Cells verify exported certificate size | C# compare original pfx size with signed workbook certificate | digital signature certificate length mismatch Aspose.Cells | how to validate certificate byte count after signing Excel file
// Developer Intent: Ensure that the certificate extracted from a signed workbook has the identical byte length as the original .pfx file.
// Use Cases: Sign an Excel workbook with a .pfx certificate using Aspose.Cells. | Persist the signed workbook to a MemoryStream and reload it for verification. | Export the embedded certificate from the loaded workbook and compare its size to the source file. | Log or handle cases where the certificate size does not match, indicating potential corruption.
// AI Prompts: Write C# code that signs an Aspose.Cells workbook with a .pfx certificate and confirms the exported certificate size equals the original file size. | Create a reusable function that takes a certificate path and a Workbook, applies a digital signature, reloads the workbook, and returns true if the certificate byte count matches. | Explain steps to troubleshoot a size mismatch when validating an exported certificate from a signed Excel file using Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsCertificateValidation
{
    // Loads a .pfx certificate, records its byte length, signs an Aspose.Cells workbook, saves it to a MemoryStream, reloads the file, extracts the embedded certificate, and checks that the exported certificate size matches the original file size.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the original certificate file (e.g., .pfx)
                string certPath = "certificate.pfx";

                // Ensure the certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the original certificate bytes and get its size
                byte[] originalCertBytes = File.ReadAllBytes(certPath);
                long originalSize = originalCertBytes.Length;
                Console.WriteLine($"Original certificate size: {originalSize} bytes");

                // Create a new workbook and add some data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Demo for certificate validation");

                // Create a digital signature collection and add a signature using the certificate
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                DigitalSignature signature = new DigitalSignature(
                    originalCertBytes,   // certificate data
                    "certPassword",      // certificate password (if any)
                    "Demo Signature",    // signature comment
                    DateTime.Now);       // signing time
                signatures.Add(signature);

                // Apply the digital signature to the workbook
                workbook.SetDigitalSignature(signatures);

                // Save to a memory stream and reload to verify the signature
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0; // reset stream position for reading

                    // Load the workbook back from the memory stream
                    Workbook loadedWorkbook = new Workbook(ms);

                    // Retrieve the digital signatures from the loaded workbook
                    DigitalSignatureCollection loadedSignatures = loadedWorkbook.GetDigitalSignature();

                    // Iterate through signatures and compare certificate sizes
                    foreach (DigitalSignature loadedSignature in loadedSignatures)
                    {
                        // Export the certificate from the loaded signature (PKCS#12 format)
                        byte[] exportedCertBytes = loadedSignature.Certificate?.Export(X509ContentType.Pkcs12) ?? Array.Empty<byte>();
                        long exportedSize = exportedCertBytes.Length;
                        Console.WriteLine($"Exported certificate size: {exportedSize} bytes");

                        // Validate that the exported size matches the original size
                        bool isSizeMatch = exportedSize == originalSize;
                        Console.WriteLine($"Size match: {isSizeMatch}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
