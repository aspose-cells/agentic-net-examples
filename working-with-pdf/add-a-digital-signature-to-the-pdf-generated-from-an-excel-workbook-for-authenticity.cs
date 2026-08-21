// Title: C# – Add an X509 digital signature to a PDF created from an Excel workbook with Aspose.Cells
// Description: Demonstrates how to load an X509Certificate2 (.pfx), create a DigitalSignature, attach it to a Workbook, and export the signed workbook as a PDF using Aspose.Cells for .NET, ensuring document authenticity and compliance.
// Keywords: Aspose.Cells digital signature PDF | C# sign PDF from Excel | X509Certificate2 Aspose.Cells | Workbook.SetDigitalSignature example | PDF authenticity Aspose | convert signed workbook to PDF
// Common Searches: how to digitally sign a PDF generated from Excel using Aspose.Cells | Aspose.Cells add X509 certificate before saving as PDF | C# code to embed digital signature in PDF with Aspose | sign Excel workbook and export to signed PDF | Aspose.Cells PDF signing tutorial
// Developer Intent: Apply an X509 digital signature to an Excel workbook and produce a signed PDF with Aspose.Cells for .NET.
// Use Cases: Secure financial statements before distribution to meet regulatory standards. | Authenticate contract PDFs generated from spreadsheet data. | Automate batch signing of invoices exported from Excel.
// AI Prompts: Write C# code that loads a .pfx certificate, creates a DigitalSignature, attaches it to an Aspose.Cells Workbook, and saves the result as a signed PDF. | Explain how to verify the digital signature in a PDF produced by Aspose.Cells after conversion. | Provide a loop example that signs multiple workbooks and generates corresponding signed PDFs using Aspose.Cells.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Demonstrates how to load an X509Certificate2 (.pfx), create a DigitalSignature, attach it to a Workbook, and export the signed workbook as a PDF using Aspose.Cells for .NET, ensuring document authenticity and compliance.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Document signed before PDF conversion");

        // Load the signing certificate (replace with your own .pfx file and password)
        X509Certificate2 certificate = new X509Certificate2("mycert.pfx", "password");

        // Create a digital signature using the certificate
        DigitalSignature digitalSignature = new DigitalSignature(
            certificate,               // certificate containing private key
            "PDF Generation Signature", // comment/description
            DateTime.Now);              // signing time

        // Add the signature to a collection and attach it to the workbook
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(digitalSignature);
        workbook.SetDigitalSignature(signatureCollection);

        // Convert the signed workbook to PDF
        workbook.Save("SignedDocument.pdf", SaveFormat.Pdf);
    }
}
