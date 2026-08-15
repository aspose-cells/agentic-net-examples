// Title: Verify Exported X509 Certificate Size Using Aspose.Cells (C#)
// Description: Loads an Excel workbook, extracts the first digital signature's X509Certificate2, exports it to a .cer file, and confirms that the file size matches the original certificate byte array.
// Keywords: Aspose.Cells export certificate | C# digital signature size check | X509Certificate2 file length verification | Excel workbook certificate export | compare .cer file size
// Common Searches: Aspose.Cells export X509 certificate and verify size | C# compare exported .cer file length with certificate bytes | how to check certificate file size after export from Excel | validate digital signature certificate size using Aspose.Cells
// Developer Intent: Ensure the .cer file created from a workbook's digital signature has the same byte length as the original certificate data.
// Use Cases: Automated integrity check of exported certificates before transmission to external services. | CI/CD validation step that flags corrupted or incomplete certificate exports. | Logging and alerting when file‑system issues cause size mismatches during export.
// AI Prompts: Create a reusable C# method that loads a workbook, extracts the first digital signature certificate with Aspose.Cells, writes it to a .cer file, and returns true if the file size equals the certificate byte array length. | Add comprehensive error handling and structured logging to each step of the certificate export and size verification process. | Write unit tests that mock a signed Workbook, export the certificate, and verify both matching and mismatching size scenarios.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsCertificateSizeCheck
{
    // Loads an Excel workbook, extracts the first digital signature's X509Certificate2, exports it to a .cer file, and confirms that the file size matches the original certificate byte array.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Path to the workbook that contains a digital signature
                string workbookPath = "SignedWorkbook.xlsx";

                // Ensure the workbook file exists before loading
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Retrieve the digital signatures collection
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                // Get the first signature if any exist
                DigitalSignature signature = null;
                if (signatures != null)
                {
                    foreach (DigitalSignature ds in signatures)
                    {
                        signature = ds;
                        break; // only need the first one
                    }
                }

                if (signature == null)
                {
                    Console.WriteLine("No digital signatures found in the workbook.");
                    return;
                }

                // Get the X509Certificate2 object used for signing
                X509Certificate2 cert = signature.Certificate;

                // Export the certificate to a byte array (raw certificate data)
                byte[] certBytes = cert.Export(X509ContentType.Cert);

                // Define the output file for the exported certificate
                string exportedCertPath = "ExportedCertificate.cer";

                // Write the certificate bytes to the file
                try
                {
                    File.WriteAllBytes(exportedCertPath, certBytes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write certificate file: {ex.Message}");
                    return;
                }

                // Get the file size in bytes
                long fileSize = new FileInfo(exportedCertPath).Length;

                // Compare the file size with the original byte array length
                bool sizeMatches = fileSize == certBytes.Length;

                Console.WriteLine($"Exported certificate file size: {fileSize} bytes");
                Console.WriteLine($"Original certificate byte array length: {certBytes.Length} bytes");
                Console.WriteLine($"Size matches expected length: {sizeMatches}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
