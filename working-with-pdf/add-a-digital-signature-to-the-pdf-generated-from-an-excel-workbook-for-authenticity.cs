// Title: C# – Add a Digital Signature to a PDF Exported from an Excel Workbook with Aspose.Cells
// Description: Learn how to create an Excel workbook, attach an X509Certificate2 digital signature, and export the signed workbook as a PDF using Aspose.Cells for .NET. The example also shows how to verify the signature before saving.
// Keywords: Aspose.Cells digital signature PDF | C# sign PDF from Excel | X509Certificate2 Aspose.Cells | Workbook.SetDigitalSignature | Export signed Excel to PDF | IsDigitallySigned property | PDF authenticity .NET
// Common Searches: how to digitally sign a PDF generated from Excel using Aspose.Cells | C# code for adding X509 certificate signature to Aspose.Cells PDF export | verify workbook digital signature before PDF conversion Aspose.Cells | add multiple digital signatures to an Excel workbook and save as PDF .NET | Aspose.Cells PDF signing example
// Developer Intent: Apply an X509 certificate‑based digital signature to an Excel workbook and produce a signed PDF with Aspose.Cells.
// Use Cases: Securely sign financial statements before distributing them as PDFs. | Automate QA approval stamps on nightly report PDFs generated from Excel. | Meet regulatory requirements by confirming a workbook is signed prior to PDF export.
// AI Prompts: Generate C# code that loads a .pfx certificate and adds a digital signature to an Aspose.Cells workbook before saving it as PDF. | Show how to check the IsDigitallySigned flag after attaching a DigitalSignatureCollection and before PDF conversion. | Provide an example of adding several DigitalSignature objects with different reasons to a workbook and exporting the signed document as PDF.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Rendering; // For SaveFormat

// Learn how to create an Excel workbook, attach an X509Certificate2 digital signature, and export the signed workbook as a PDF using Aspose.Cells for .NET. The example also shows how to verify the signature before saving.
class AddDigitalSignatureToPdf
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Document requiring digital signature");

        // Load the signing certificate (replace with your own .pfx file and password)
        X509Certificate2 certificate = new X509Certificate2("myCertificate.pfx", "certPassword");

        // Create a digital signature instance
        DigitalSignature signature = new DigitalSignature(certificate, "Approved by QA", DateTime.UtcNow);

        // Add the signature to a collection and attach it to the workbook
        DigitalSignatureCollection signatures = new DigitalSignatureCollection();
        signatures.Add(signature);
        workbook.SetDigitalSignature(signatures); // or workbook.AddDigitalSignature(signatures);

        // Save the signed workbook as a PDF file
        workbook.Save("SignedDocument.pdf", SaveFormat.Pdf);

        // Optional: verify that the workbook is digitally signed before conversion
        Console.WriteLine("Workbook digitally signed: " + workbook.IsDigitallySigned);
    }
}
