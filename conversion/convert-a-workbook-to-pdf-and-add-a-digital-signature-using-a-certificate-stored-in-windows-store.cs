// Title: Sign an Excel workbook with a Windows certificate store and convert it to PDF using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file, locate an X509Certificate2 in the Current User Personal store, create a DigitalSignature, attach it to the workbook, and save both the signed .xlsx and a PDF using Aspose.Cells. | Search the Windows certificate store by subject or thumbprint, build a DigitalSignatureCollection, add it to a Workbook object, then export the signed workbook to PDF in C#. | Apply a Windows store certificate to an Excel file with Aspose.Cells' DigitalSignature class and generate a signed PDF using SaveFormat.Pdf.
// Common Searches: Aspose.Cells sign Excel file with certificate from Windows store and export to PDF C# | How to add a digital signature to an .xlsx using a Windows certificate and then convert to PDF in .NET | C# retrieve X509Certificate2 from Current User store for Aspose.Cells workbook signing
// Tags: Aspose.Cells digital signature from Windows store | convert signed Excel to PDF using Aspose.Cells | C# X509Certificate2 workbook signing | SaveFormat.Pdf with digital signature Aspose | DigitalSignatureCollection usage Aspose.Cells

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// The example loads an Excel workbook, retrieves an X509 certificate from the Current User Personal store, creates a DigitalSignature, adds it to the workbook, saves the signed workbook as .xlsx, and then converts the signed workbook to PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Open the Current User's Personal (My) certificate store
        X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);

        // Locate the certificate you want to use.
        // Adjust the search condition (e.g., Subject, Thumbprint) to match your certificate.
        X509Certificate2 certificate = null;
        foreach (X509Certificate2 cert in store.Certificates)
        {
            if (cert.Subject.Contains("YourCompany")) // <-- replace with appropriate identifier
            {
                certificate = cert;
                break;
            }
        }

        store.Close();

        if (certificate == null)
        {
            Console.WriteLine("Certificate not found in the Windows store.");
            return;
        }

        // Create a digital signature using the found certificate
        DigitalSignature digitalSignature = new DigitalSignature(
            certificate,
            "Workbook signed with Windows store certificate",
            DateTime.UtcNow);

        // Add the signature to a collection and attach it to the workbook
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(digitalSignature);
        workbook.AddDigitalSignature(signatureCollection);

        // Optionally save the signed workbook as an Excel file
        workbook.Save("SignedWorkbook.xlsx");

        // Convert the signed workbook to PDF
        workbook.Save("SignedWorkbook.pdf", SaveFormat.Pdf);

        Console.WriteLine("Workbook signed and converted to PDF successfully.");
    }
}
