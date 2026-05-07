using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing signed workbook
            string signedWorkbookPath = "SignedWorkbook.xlsx";

            // Load the workbook (preserves existing digital signatures)
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Load the certificate used for the new signature
            // Replace with your actual certificate file path and password
            string certificatePath = "mycert.pfx";
            string certificatePassword = "password";
            X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword);

            // Create a new digital signature
            DigitalSignature additionalSignature = new DigitalSignature(
                certificate,
                "Additional signature added on " + DateTime.Now,
                DateTime.Now);

            // Prepare a collection containing the new signature
            DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
            signatureCollection.Add(additionalSignature);

            // Add the new signature to the workbook without removing existing ones
            workbook.AddDigitalSignature(signatureCollection);

            // Save the workbook with the added signature
            string outputPath = "SignedWorkbook_AdditionalSignature.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Additional digital signature applied successfully.");
        }
    }
}