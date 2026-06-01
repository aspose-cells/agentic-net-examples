using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class CertificateThumbprintVerificationDemo
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
            // Path to the certificate (PFX) and its password
            string certPath = "mycert.pfx";
            string certPassword = "certPassword";

            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
                return;
            }

            // Load the original certificate that will be used for signing
            X509Certificate2 originalCertificate = new X509Certificate2(certPath, certPassword);
            string originalThumbprint = originalCertificate.Thumbprint;
            Console.WriteLine($"Original Certificate Thumbprint: {originalThumbprint}");

            // -------------------------------------------------
            // Create a new workbook and add some sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // create
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Document to be digitally signed");

            // -------------------------------------------------
            // Create a digital signature using the original certificate
            // -------------------------------------------------
            DigitalSignature signature = new DigitalSignature(originalCertificate, "Signed by Aspose", DateTime.Now);
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // Add the digital signature to the workbook
            workbook.AddDigitalSignature(signatures); // add digital signature

            // Save the signed workbook
            string signedPath = "SignedWorkbook.xlsx";
            workbook.Save(signedPath); // save
            Console.WriteLine($"Signed workbook saved to: {signedPath}");

            // -------------------------------------------------
            // Load the signed workbook and verify the certificate thumbprint
            // -------------------------------------------------
            if (!File.Exists(signedPath))
            {
                Console.WriteLine($"Signed workbook not found: {signedPath}");
                return;
            }

            Workbook signedWorkbook = new Workbook(signedPath); // load
            DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

            foreach (DigitalSignature loadedSignature in loadedSignatures)
            {
                // Get the certificate from the loaded signature
                X509Certificate2 loadedCertificate = loadedSignature.Certificate;
                string loadedThumbprint = loadedCertificate?.Thumbprint;

                Console.WriteLine($"Loaded Signature Thumbprint: {loadedThumbprint}");

                // Compare thumbprints
                bool isThumbprintMatch = string.Equals(originalThumbprint, loadedThumbprint, StringComparison.OrdinalIgnoreCase);
                Console.WriteLine($"Thumbprint match: {isThumbprintMatch}");
            }
        }
    }
}