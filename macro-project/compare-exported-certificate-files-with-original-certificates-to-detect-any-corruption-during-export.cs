using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

class CompareCertificates
{
    static void Main()
    {
        try
        {
            // Paths to the original certificate and the signed workbook
            string originalCertPath = "original.pfx";
            string certPassword = "password";
            string signedWorkbookPath = "signed.xlsx";

            // Verify that the certificate file exists
            if (!File.Exists(originalCertPath))
            {
                Console.WriteLine($"Certificate file not found: {originalCertPath}");
                return;
            }

            // Load the original certificate (including private key)
            X509Certificate2 originalCert = new X509Certificate2(originalCertPath, certPassword);
            // Export the original certificate to a byte array for later comparison
            byte[] originalRawData = originalCert.Export(X509ContentType.Pkcs12);

            // Create a new workbook and add some content
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Signed Workbook");

            // Create a digital signature using the original certificate
            DigitalSignature signature = new DigitalSignature(originalCert, "Demo Signature", DateTime.Now);

            // Add the digital signature to a collection and attach it to the workbook
            DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
            signatureCollection.Add(signature);
            workbook.AddDigitalSignature(signatureCollection);

            // Save the signed workbook
            workbook.Save(signedWorkbookPath, SaveFormat.Xlsx);

            // Verify that the signed workbook was created
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Failed to create signed workbook: {signedWorkbookPath}");
                return;
            }

            // Load the signed workbook and retrieve its digital signatures
            Workbook signedWorkbook = new Workbook(signedWorkbookPath);
            DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

            // Compare each exported certificate with the original one
            foreach (DigitalSignature loadedSignature in loadedSignatures)
            {
                X509Certificate2 exportedCert = loadedSignature.Certificate;
                byte[] exportedRawData = exportedCert.Export(X509ContentType.Pkcs12);

                bool certificatesMatch = AreByteArraysEqual(originalRawData, exportedRawData);
                Console.WriteLine($"Certificate match: {certificatesMatch}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to compare two byte arrays
    static bool AreByteArraysEqual(byte[] a, byte[] b)
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