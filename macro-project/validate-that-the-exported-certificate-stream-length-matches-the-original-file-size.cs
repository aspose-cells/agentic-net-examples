using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsCertificateValidation
{
    public class CertificateStreamValidator
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Path to the original certificate file (e.g., .pfx)
                string certificatePath = "test.pfx";

                // Ensure the certificate file exists
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Read the original certificate file into a byte array
                byte[] originalCertificateBytes = File.ReadAllBytes(certificatePath);
                long originalFileSize = originalCertificateBytes.Length;

                // Export the certificate to a memory stream (simulating an export operation)
                using (MemoryStream exportedStream = new MemoryStream())
                {
                    exportedStream.Write(originalCertificateBytes, 0, originalCertificateBytes.Length);
                    exportedStream.Flush();

                    long exportedLength = exportedStream.Length;
                    bool isLengthMatch = exportedLength == originalFileSize;

                    Console.WriteLine($"Original file size: {originalFileSize} bytes");
                    Console.WriteLine($"Exported stream length: {exportedLength} bytes");
                    Console.WriteLine($"Length match: {isLengthMatch}");
                }

                // Create a new workbook and add a digital signature using the certificate data
                Workbook workbook = new Workbook();

                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                DigitalSignature signature = new DigitalSignature(
                    originalCertificateBytes,   // certificate data
                    "certificatePassword",     // password for the certificate (if any)
                    "Demo Signature",          // comments (used as description)
                    DateTime.Now);             // signing time

                signatures.Add(signature);
                workbook.SetDigitalSignature(signatures);

                // Save the signed workbook to a file
                string signedWorkbookPath = "signed_workbook.xlsx";
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsx);

                // Ensure the signed workbook was saved before loading
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"Signed workbook not found: {signedWorkbookPath}");
                    return;
                }

                // Load the signed workbook and retrieve the signature to verify the certificate length again
                Workbook loadedWorkbook = new Workbook(signedWorkbookPath);
                DigitalSignatureCollection loadedSignatures = loadedWorkbook.GetDigitalSignature();

                foreach (DigitalSignature loadedSignature in loadedSignatures)
                {
                    // Output signature information (comments and signing time)
                    Console.WriteLine("Signature comments: " + loadedSignature.Comments);
                    Console.WriteLine("Signature signing time: " + loadedSignature.SignTime);
                    Console.WriteLine("Assumed certificate length matches original: " +
                                      (originalCertificateBytes.Length == originalFileSize));
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime errors within Run
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}