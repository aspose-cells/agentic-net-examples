using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Define the subject name to search for in the Windows certificate store
            const string certificateSubjectName = "CN=MyCodeSigningCert";

            // Open the Current User's Personal (My) certificate store
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);

            // Find a certificate that matches the subject name and has a private key (required for signing)
            X509Certificate2 signingCertificate = store.Certificates
                .Cast<X509Certificate2>()
                .FirstOrDefault(cert =>
                    cert.Subject.Equals(certificateSubjectName, StringComparison.OrdinalIgnoreCase) &&
                    cert.HasPrivateKey);

            store.Close();

            if (signingCertificate == null)
            {
                Console.WriteLine("Certificate with subject '{0}' not found in the store.", certificateSubjectName);
                return;
            }

            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a new workbook
            // Example: add some content
            workbook.Worksheets[0].Cells["A1"].PutValue("Digitally Signed Workbook");

            // Create a digital signature using the certificate, a comment, and the current UTC time
            DigitalSignature signature = new DigitalSignature(
                signingCertificate,
                "Code signing of workbook",
                DateTime.UtcNow);

            // Add the signature to a collection
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // Apply the digital signature to the workbook
            workbook.SetDigitalSignature(signatures);

            // Save the signed workbook
            string outputPath = "SignedWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Workbook signed and saved to: " + outputPath);
        }
    }
}