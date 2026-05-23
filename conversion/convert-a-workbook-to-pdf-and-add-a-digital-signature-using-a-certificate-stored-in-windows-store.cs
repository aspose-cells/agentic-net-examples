using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Rendering; // For SaveFormat enum

class WorkbookToPdfWithSignature
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path for the resulting PDF file
        string pdfOutputPath = "output.pdf";

        // Load the workbook from file
        Workbook workbook = new Workbook(sourcePath);

        // Open the current user's personal certificate store
        X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);

        // Find a certificate by its subject name (replace with your own criteria)
        X509Certificate2 certificate = store.Certificates
            .Find(X509FindType.FindBySubjectName, "YourCertificateSubject", validOnly: false)
            .OfType<X509Certificate2>()
            .FirstOrDefault();

        store.Close();

        if (certificate == null)
        {
            Console.WriteLine("Certificate not found in the Windows store.");
            return;
        }

        // Create a digital signature using the certificate
        DigitalSignature signature = new DigitalSignature(
            certificate,
            "Signed by Aspose.Cells example",
            DateTime.Now);

        // Add the signature to a collection
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(signature);

        // Attach the digital signature to the workbook
        workbook.AddDigitalSignature(signatureCollection);

        // Save the signed workbook as a PDF document
        workbook.Save(pdfOutputPath, SaveFormat.Pdf);

        Console.WriteLine("Workbook converted to PDF and digitally signed successfully.");
    }
}