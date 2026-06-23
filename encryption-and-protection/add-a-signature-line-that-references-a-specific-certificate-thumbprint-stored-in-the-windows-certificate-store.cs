using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Load the certificate from the Current User's Personal store using its thumbprint
            const string thumbprint = "YOUR_CERT_THUMBPRINT"; // replace with actual thumbprint (no spaces)
            X509Certificate2 certificate = null;

            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certs.Count > 0)
                {
                    certificate = certs[0];
                }
                store.Close();
            }

            if (certificate == null)
            {
                Console.WriteLine("Certificate with the specified thumbprint was not found.");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a signature line and assign a unique Id (to link with DigitalSignature later)
            SignatureLine sigLine = new SignatureLine
            {
                Signer = "John Doe",
                Title = "Approver",
                Email = "john.doe@example.com",
                IsLine = true,
                AllowComments = true,
                ShowSignedDate = true,
                Instructions = "Please sign to approve.",
                Id = Guid.NewGuid(),               // unique identifier for the line
                SignatureLineType = SignatureType.Custom,
                ProviderId = Guid.NewGuid()        // custom provider id (optional)
            };

            // Add the signature line to the worksheet as a picture (required for Excel UI)
            Picture picture = sheet.Shapes.AddSignatureLine(2, 2, sigLine);

            // Create a digital signature using the loaded certificate
            DigitalSignature digitalSignature = new DigitalSignature(certificate, "Document approved", DateTime.Now)
            {
                // Link the digital signature to the signature line via the same Id
                Id = sigLine.Id
            };

            // Add the digital signature to a collection and apply it to the workbook
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(digitalSignature);
            workbook.SetDigitalSignature(signatures);

            // Save the signed workbook
            workbook.Save("SignedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}