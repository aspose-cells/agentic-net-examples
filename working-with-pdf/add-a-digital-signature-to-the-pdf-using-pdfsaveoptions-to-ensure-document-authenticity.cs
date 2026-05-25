using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfDigitalSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Document with PDF digital signature");

            // Load a certificate (replace with your actual .pfx file path and password)
            X509Certificate2 certificate = new X509Certificate2("myCertificate.pfx", "certPassword");

            // Create a digital signature for the workbook (optional, signs the Excel content)
            DigitalSignature signature = new DigitalSignature(certificate, "PDF Signed Document", DateTime.UtcNow);
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);
            workbook.SetDigitalSignature(signatures);

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Configure PDF security options (acts as a digital signature for PDF authenticity)
            PdfSecurityOptions pdfSecurity = new PdfSecurityOptions
            {
                // Owner password allows full control over the PDF
                OwnerPassword = "OwnerSecret123",
                // User password is required to open the PDF
                UserPassword = "UserSecret123",
                // Permissions can be set as needed
                PrintPermission = true,
                FullQualityPrintPermission = true,
                ModifyDocumentPermission = false,
                ExtractContentPermission = false
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = pdfSecurity;

            // Save the workbook as a PDF with the configured security (digital signature) options
            workbook.Save("SignedDocument.pdf", pdfSaveOptions);

            Console.WriteLine("PDF saved with security (digital signature) options.");
        }
    }
}