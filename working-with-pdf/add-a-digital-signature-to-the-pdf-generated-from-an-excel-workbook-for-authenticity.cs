using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Rendering; // For PdfSaveOptions

class AddDigitalSignatureToPdf
{
    static void Main()
    {
        // 1. Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Document with Digital Signature");

        // 2. Load the signing certificate (replace with your own .pfx file and password)
        X509Certificate2 certificate = new X509Certificate2("myCertificate.pfx", "certPassword");

        // 3. Create a digital signature using the certificate
        DigitalSignature signature = new DigitalSignature(certificate, "Approved by QA", DateTime.Now);

        // 4. Add the signature to a collection
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(signature);

        // 5. Attach the digital signature to the workbook
        workbook.AddDigitalSignature(signatureCollection);

        // 6. Prepare PDF save options (optional: you can set security options here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // 7. Save the signed workbook as a PDF file
        workbook.Save("SignedDocument.pdf", pdfOptions);

        // Cleanup
        workbook.Dispose();
    }
}