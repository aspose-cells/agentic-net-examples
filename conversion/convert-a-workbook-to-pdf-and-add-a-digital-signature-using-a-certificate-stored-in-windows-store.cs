// Title: Sign an Aspose.Cells Workbook and Convert to PDF using a Windows Certificate Store (C#)
// Description: Demonstrates how to load an X509Certificate2 from the CurrentUser Personal store, create an Aspose.Cells DigitalSignature, attach it to a workbook, and save the signed workbook as a PDF file.
// Keywords: Aspose.Cells PDF conversion | digital signature C# | Windows certificate store | X509Certificate2 load | AddDigitalSignature | .NET Excel to PDF | certificate subject lookup | Aspose.Cells DigitalSignatureCollection
// Common Searches: Aspose.Cells sign PDF with Windows certificate | C# load X509Certificate2 from store for Aspose.Cells | Add digital signature to Excel workbook before PDF export | Convert signed workbook to PDF using Aspose.Cells | Retrieve certificate by subject name in .NET
// Developer Intent: Generate a PDF from an Excel workbook and embed a digital signature sourced from the Windows certificate store.
// Use Cases: Create compliance‑ready PDF reports that are automatically signed with a user‑specific certificate. | Batch‑process multiple Excel files, applying a distinct store‑based certificate to each before PDF conversion. | Expose a REST endpoint that receives Excel data, signs it with a server‑side certificate, and returns a signed PDF.
// AI Prompts: Write C# code to fetch an X509Certificate2 from the CurrentUser Personal store by subject name and use it to sign an Aspose.Cells workbook before saving as PDF. | Explain the steps for attaching a digital signature to an Aspose.Cells workbook and exporting it to a signed PDF file. | Provide a tutorial on using Aspose.Cells.DigitalSignatures with certificates stored in Windows for PDF generation.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Demonstrates how to load an X509Certificate2 from the CurrentUser Personal store, create an Aspose.Cells DigitalSignature, attach it to a workbook, and save the signed workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Signed PDF Example");

        // Load a certificate from the current user's Personal store
        X509Certificate2 certificate = LoadCertificateFromStore("MyCertificateSubject");

        if (certificate == null)
        {
            Console.WriteLine("Certificate not found in the store.");
            return;
        }

        // Create a digital signature using the loaded certificate
        DigitalSignature signature = new DigitalSignature(certificate, "Workbook signed", DateTime.Now);

        // Add the signature to a collection and attach it to the workbook
        DigitalSignatureCollection signatures = new DigitalSignatureCollection();
        signatures.Add(signature);
        workbook.AddDigitalSignature(signatures);

        // Convert the signed workbook to PDF
        workbook.Save("SignedWorkbook.pdf", SaveFormat.Pdf);
    }

    // Helper method to retrieve a certificate by subject name from the Windows certificate store
    static X509Certificate2 LoadCertificateFromStore(string subjectContains)
    {
        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);
            foreach (X509Certificate2 cert in store.Certificates)
            {
                if (cert.Subject.IndexOf(subjectContains, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return cert;
                }
            }
        }
        return null;
    }
}
