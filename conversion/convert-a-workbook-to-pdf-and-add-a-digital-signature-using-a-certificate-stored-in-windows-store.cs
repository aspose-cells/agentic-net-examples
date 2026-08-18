// Title: C# – Convert Aspose.Cells Workbook to PDF and Apply a Windows Store Digital Signature
// Description: Shows how to load or create an Aspose.Cells workbook, locate an X509Certificate2 in the CurrentUser Windows certificate store by thumbprint, build a DigitalSignature, attach it to the workbook, and save the signed document as a PDF.
// Keywords: Aspose.Cells | C# PDF conversion | digital signature | Windows certificate store | X509Certificate2 | SaveFormat.Pdf | AddDigitalSignature | Workbook to PDF | certificate thumbprint | signed PDF generation
// Common Searches: Aspose.Cells sign PDF with Windows certificate | C# convert Excel to PDF and add digital signature | retrieve X509Certificate2 by thumbprint C# | add digital signature to workbook before PDF export | Aspose.Cells digital signature example
// Developer Intent: Generate a PDF from an Excel workbook and embed a digital signature using a certificate retrieved from the Windows certificate store.
// Use Cases: Produce compliance‑ready PDF reports by converting Excel workbooks and signing them with the organization’s certificate stored in Windows. | Automate creation of digitally signed invoices by converting Excel templates to PDF and applying each user’s personal certificate from the CurrentUser store. | Integrate signed PDF generation into a Windows service that fetches the signing certificate from the store and outputs signed PDFs on a schedule.
// AI Prompts: Provide C# code that loads an Aspose.Cells workbook, finds a certificate by thumbprint in the CurrentUser Windows store, creates a DigitalSignature, adds it to the workbook, and saves the result as a PDF. | Explain how to handle a missing certificate in the Windows store and suggest fallback strategies for Aspose.Cells PDF signing in C#. | Show how to attach multiple digital signatures to a workbook before exporting it to PDF using Aspose.Cells.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Rendering; // For SaveFormat enum

// Shows how to load or create an Aspose.Cells workbook, locate an X509Certificate2 in the CurrentUser Windows certificate store by thumbprint, build a DigitalSignature, attach it to the workbook, and save the signed document as a PDF.
class WorkbookToPdfWithSignature
{
    static void Main()
    {
        // 1. Create or load a workbook
        Workbook workbook = new Workbook(); // creates a new workbook
        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Signed PDF Example");

        // 2. Retrieve a certificate from the Windows certificate store
        //    (e.g., a certificate with a specific thumbprint)
        string thumbprint = "YOUR_CERTIFICATE_THUMBPRINT".Replace(" ", "").ToUpperInvariant(); // replace with actual thumbprint
        X509Certificate2 certificate = null;

        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);
            foreach (var cert in store.Certificates)
            {
                if (cert.Thumbprint != null && cert.Thumbprint.Equals(thumbprint, StringComparison.OrdinalIgnoreCase))
                {
                    certificate = cert;
                    break;
                }
            }
            store.Close();
        }

        if (certificate == null)
        {
            Console.WriteLine("Certificate not found in the Windows store.");
            return;
        }

        // 3. Create a digital signature using the retrieved certificate
        DigitalSignature signature = new DigitalSignature(certificate, "Workbook signed for PDF conversion", DateTime.Now);

        // 4. Add the signature to the workbook
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(signature);
        workbook.AddDigitalSignature(signatureCollection); // adds the digital signature to the OOXML workbook

        // 5. Save the workbook as a PDF file
        string pdfPath = "SignedWorkbook.pdf";
        workbook.Save(pdfPath, SaveFormat.Pdf);

        Console.WriteLine($"Workbook has been signed and saved as PDF to: {pdfPath}");
    }
}
