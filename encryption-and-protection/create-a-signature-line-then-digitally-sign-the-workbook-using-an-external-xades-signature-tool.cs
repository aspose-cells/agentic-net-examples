using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.DigitalSignatures;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure a signature line that will appear in the worksheet
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";
        signatureLine.Email = "john.doe@example.com";
        signatureLine.Instructions = "Please sign to approve the document.";
        signatureLine.AllowComments = true;
        signatureLine.ShowSignedDate = true;
        signatureLine.IsLine = true;
        signatureLine.Id = Guid.NewGuid(); // unique identifier for linking

        // Add the signature line picture to the worksheet at row 5, column 2 (zero‑based indices)
        Picture picture = worksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Load the signing certificate (replace with actual path and password)
        string certPath = "mycert.pfx";
        string certPassword = "password";
        X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

        // Create a digital signature using the certificate and set XAdES type
        DigitalSignature digitalSignature = new DigitalSignature(certificate, "Approved by John Doe", DateTime.UtcNow);
        digitalSignature.XAdESType = XAdESType.XAdES; // enable XAdES support
        digitalSignature.Id = signatureLine.Id; // link signature line with digital signature

        // Add the digital signature to a collection and apply it to the workbook
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(digitalSignature);
        workbook.SetDigitalSignature(signatureCollection);

        // Save the workbook with the embedded signature line and digital signature
        workbook.Save("SignedWorkbook.xlsx");
    }
}