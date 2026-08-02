// Title: Verify exported certificate from an Aspose.Cells signed workbook matches the original PFX (C#)
// Description: Loads a PFX, signs a workbook with Aspose.Cells, saves it, reloads the file, extracts the embedded certificate, exports it as PKCS#12, and compares the byte arrays to detect any corruption.
// Keywords: Aspose.Cells | C# | digital signature | certificate export | PFX verification | X509Certificate2 | byte array comparison | Excel workbook signing | PKCS#12 | certificate integrity
// Common Searches: Aspose.Cells verify exported certificate | compare original PFX with certificate from signed Excel | detect certificate corruption after Aspose.Cells signing | C# extract digital signature certificate from .xlsx | how to compare X509Certificate2 byte arrays
// Developer Intent: Ensure the certificate embedded in a signed Excel file is identical to the source PFX, confirming no corruption during export.
// Use Cases: Automated CI validation of signed workbooks by extracting and comparing certificates. | Security audit to detect tampering of Excel digital signatures. | Debugging certificate export issues in Aspose.Cells applications.
// AI Prompts: Write C# code that opens a signed .xlsx created with Aspose.Cells, extracts its digital signature certificate, and validates it against a given .pfx file using byte‑wise comparison. | Explain how to handle password‑protected PFX files when exporting certificates from Aspose.Cells signatures and how to report mismatches. | Suggest enhancements such as hashing, detailed diff, and logging for certificate integrity checks in Aspose.Cells workflows.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Loads a PFX, signs a workbook with Aspose.Cells, saves it, reloads the file, extracts the embedded certificate, exports it as PKCS#12, and compares the byte arrays to detect any corruption.
class CertificateExportComparison
{
    static void Main()
    {
        try
        {
            // Paths to the original certificate and the signed workbook
            string certPath = "original.pfx";
            string certPassword = "password";
            string signedFile = "signed.xlsx";

            // Verify that the certificate file exists
            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
                return;
            }

            // Load the original certificate (including private key)
            X509Certificate2 originalCert = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.Exportable);

            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Certificate Export Test");

            // Create a digital signature using the original certificate
            DigitalSignature signature = new DigitalSignature(originalCert, "Test Signature", DateTime.UtcNow);
            DigitalSignatureCollection collection = new DigitalSignatureCollection();
            collection.Add(signature);

            // Add the digital signature to the workbook
            wb.AddDigitalSignature(collection);

            // Save the signed workbook
            wb.Save(signedFile, SaveFormat.Xlsx);

            // Verify that the signed workbook was saved
            if (!File.Exists(signedFile))
            {
                Console.WriteLine($"Failed to create signed workbook: {signedFile}");
                return;
            }

            // Load the signed workbook for verification
            Workbook signedWb = new Workbook(signedFile);
            Console.WriteLine("Workbook signed: " + signedWb.IsDigitallySigned);

            // Retrieve the digital signatures from the loaded workbook
            DigitalSignatureCollection loadedSignatures = signedWb.GetDigitalSignature();

            // Compare each exported certificate with the original certificate
            foreach (DigitalSignature loadedSig in loadedSignatures)
            {
                // Export the certificate from the loaded signature (PKCS#12 format)
                byte[] exportedRaw = loadedSig.Certificate.Export(X509ContentType.Pkcs12);

                // Export the original certificate in the same format
                byte[] originalRaw = originalCert.Export(X509ContentType.Pkcs12);

                // Compare the raw byte arrays
                bool isMatch = CompareByteArrays(originalRaw, exportedRaw);
                Console.WriteLine("Certificate match after export: " + isMatch);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Helper method to compare two byte arrays
    static bool CompareByteArrays(byte[] a, byte[] b)
    {
        if (a == null || b == null) return false;
        if (a.Length != b.Length) return false;
        for (int i = 0; i < a.Length; i++)
        {
            if (a[i] != b[i]) return false;
        }
        return true;
    }
}
