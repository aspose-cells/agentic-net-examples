using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

class AddAdditionalSignature
{
    static void Main()
    {
        // Path to the already signed workbook
        string sourcePath = "SignedWorkbook.xlsx";
        // Path where the workbook with the new signature will be saved
        string destPath = "SignedWorkbook_WithAdditionalSignature.xlsx";

        // Load the existing workbook (preserves current signatures)
        Workbook workbook = new Workbook(sourcePath);

        // Get the current digital signature collection; create one if none exist
        DigitalSignatureCollection signatureCollection = workbook.GetDigitalSignature() ?? new DigitalSignatureCollection();

        // Load the certificate used for signing (replace with your certificate file and password)
        X509Certificate2 certificate = new X509Certificate2("myCertificate.pfx", "certPassword");

        // Create a new digital signature
        DigitalSignature additionalSignature = new DigitalSignature(certificate, "Additional Signature", DateTime.Now);

        // Add the new signature to the existing collection
        signatureCollection.Add(additionalSignature);

        // Save the workbook; existing signatures remain valid
        workbook.Save(destPath);
    }
}