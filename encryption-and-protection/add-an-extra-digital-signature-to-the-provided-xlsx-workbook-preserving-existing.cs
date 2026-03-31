using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

class AddExtraDigitalSignature
{
    static void Main()
    {
        // Path to the existing workbook (must exist)
        string sourcePath = "input.xlsx";

        // Path where the signed workbook will be saved
        string destPath = "output_signed.xlsx";

        // Path to the PFX certificate file and its password
        string certPath = "mycert.pfx";
        string certPassword = "password";

        // Load the existing workbook (preserves all content and formatting)
        Workbook workbook = new Workbook(sourcePath);

        // Load the certificate that contains the private key
        X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

        // Create a new digital signature using the certificate
        DigitalSignature newSignature = new DigitalSignature(certificate, "Additional signature", DateTime.Now);

        // Retrieve any existing digital signatures; if none, create a new collection
        DigitalSignatureCollection signatures = workbook.GetDigitalSignature();
        if (signatures == null)
        {
            signatures = new DigitalSignatureCollection();
        }

        // Add the new signature to the collection
        signatures.Add(newSignature);

        // Add the signature collection to the workbook (preserves existing signatures)
        workbook.AddDigitalSignature(signatures);

        // Save the workbook with the added signature
        workbook.Save(destPath);
    }
}