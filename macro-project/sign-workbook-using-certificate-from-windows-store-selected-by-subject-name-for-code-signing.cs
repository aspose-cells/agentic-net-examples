using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

class SignWorkbookFromStore
{
    static void Main()
    {
        // Subject name of the code signing certificate in the Windows store
        string subjectName = "MyCodeSigningCert"; // TODO: replace with actual subject name

        // Retrieve the certificate that contains a private key
        X509Certificate2 certificate = GetCertificateBySubject(subjectName);
        if (certificate == null)
        {
            Console.WriteLine("Certificate not found.");
            return;
        }

        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Signed by Windows Store Certificate");

        // Create a digital signature using the certificate
        DigitalSignature signature = new DigitalSignature(certificate, "Code signing", DateTime.UtcNow);

        // Add the signature to a collection
        DigitalSignatureCollection signatures = new DigitalSignatureCollection();
        signatures.Add(signature);

        // Apply the digital signature to the workbook
        workbook.SetDigitalSignature(signatures);

        // Save the signed workbook
        workbook.Save("SignedWorkbook.xlsx");
        Console.WriteLine("Workbook signed and saved.");
    }

    static X509Certificate2 GetCertificateBySubject(string subjectName)
    {
        // Open the personal certificate store of the current user
        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);
            // Find certificates matching the subject name (case‑insensitive)
            X509Certificate2Collection found = store.Certificates.Find(
                X509FindType.FindBySubjectName, subjectName, validOnly: false);

            // Return the first certificate that has an associated private key
            foreach (X509Certificate2 cert in found)
            {
                if (cert.HasPrivateKey)
                    return cert;
            }
        }
        return null;
    }
}